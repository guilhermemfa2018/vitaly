VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura de Ficha de Orçamento"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   Icon            =   "frmFO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin ZEUS.chameleonButton chamCad 
      Height          =   615
      Index           =   6
      Left            =   1320
      TabIndex        =   20
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   8040
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
      MICON           =   "frmFO.frx":0CCA
      PICN            =   "frmFO.frx":0CE6
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
      TabIndex        =   21
      Tag             =   "Exporta para o Excel"
      ToolTipText     =   "Exporta para o Excel"
      Top             =   8040
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
      MICON           =   "frmFO.frx":19C0
      PICN            =   "frmFO.frx":19DC
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
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Tag             =   "Gravar dados"
      ToolTipText     =   "Gravar dados"
      Top             =   8040
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
      MICON           =   "frmFO.frx":26B6
      PICN            =   "frmFO.frx":26D2
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
      Height          =   7695
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Ficha de Orçamento"
      TabPicture(0)   =   "frmFO.frx":33AC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Itens da FO"
      TabPicture(1)   =   "frmFO.frx":33C8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SkinLabel19"
      Tab(1).Control(1)=   "chamCad(1)"
      Tab(1).Control(2)=   "chamCad(2)"
      Tab(1).Control(3)=   "chamCad(0)"
      Tab(1).Control(4)=   "Frame14"
      Tab(1).Control(5)=   "chamCad(3)"
      Tab(1).Control(6)=   "Frame17"
      Tab(1).Control(7)=   "ListView1"
      Tab(1).Control(8)=   "Text1"
      Tab(1).Control(9)=   "Frame8(0)"
      Tab(1).Control(10)=   "ScriptControl1"
      Tab(1).Control(11)=   "Shape2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Resumo da FO"
      TabPicture(2)   =   "frmFO.frx":33E4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbltotpm"
      Tab(2).Control(1)=   "lbltotl"
      Tab(2).Control(2)=   "SkinLabel39"
      Tab(2).Control(3)=   "SkinLabel38"
      Tab(2).Control(4)=   "Text4"
      Tab(2).Control(5)=   "ListView2"
      Tab(2).Control(6)=   "Frame10"
      Tab(2).Control(7)=   "Shape1"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Dados da Proposta"
      TabPicture(3)   =   "frmFO.frx":3400
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).ControlCount=   1
      Begin ACTIVESKINLibCtl.SkinLabel lbltotpm 
         Height          =   255
         Left            =   -72240
         OleObjectBlob   =   "frmFO.frx":341C
         TabIndex        =   129
         Top             =   7320
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel lbltotl 
         Height          =   255
         Left            =   -72240
         OleObjectBlob   =   "frmFO.frx":347A
         TabIndex        =   128
         Top             =   6960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
         Height          =   255
         Left            =   -74640
         OleObjectBlob   =   "frmFO.frx":34D8
         TabIndex        =   127
         Top             =   7320
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
         Height          =   255
         Left            =   -74640
         OleObjectBlob   =   "frmFO.frx":3568
         TabIndex        =   126
         Top             =   6960
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   615
         Left            =   -70320
         OleObjectBlob   =   "frmFO.frx":35EC
         TabIndex        =   106
         Top             =   3960
         Width           =   4815
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   1
         Left            =   -73680
         TabIndex        =   10
         Tag             =   "Excluir registro"
         ToolTipText     =   "Excluir registro"
         Top             =   3900
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
         MICON           =   "frmFO.frx":373C
         PICN            =   "frmFO.frx":3758
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   5040
         TabIndex        =   63
         Top             =   360
         Width           =   7575
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   3480
            TabIndex        =   65
            Top             =   2880
            Width           =   3975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   120
            TabIndex        =   66
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   3480
            TabIndex        =   67
            Top             =   2280
            Width           =   3975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   120
            TabIndex        =   68
            Top             =   2280
            Width           =   3255
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   6720
            TabIndex        =   69
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   3480
            TabIndex        =   70
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   120
            TabIndex        =   71
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   6360
            TabIndex        =   72
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   6135
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   1080
            TabIndex        =   74
            Top             =   480
            Width           =   5895
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   75
            Tag             =   "Código do Cliente"
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "frmFO.frx":4432
            TabIndex        =   117
            Top             =   2640
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":449A
            TabIndex        =   116
            Top             =   2640
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "frmFO.frx":4504
            TabIndex        =   115
            Top             =   2040
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":456A
            TabIndex        =   114
            Top             =   2040
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   6720
            OleObjectBlob   =   "frmFO.frx":45DA
            TabIndex        =   113
            Top             =   1440
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "frmFO.frx":4646
            TabIndex        =   112
            Top             =   1440
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":46B2
            TabIndex        =   111
            Top             =   1440
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   6360
            OleObjectBlob   =   "frmFO.frx":471E
            TabIndex        =   110
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":4784
            TabIndex        =   109
            Top             =   840
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "frmFO.frx":47F4
            TabIndex        =   108
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":485C
            TabIndex        =   107
            Top             =   240
            Width           =   615
         End
         Begin ZEUS.chameleonButton chameleonButton8 
            Height          =   255
            Left            =   7080
            TabIndex        =   64
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BTYPE           =   4
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
            MICON           =   "frmFO.frx":48C8
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
      Begin VB.Frame Frame6 
         Caption         =   "Dados do Contato "
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
         Left            =   5040
         TabIndex        =   57
         Top             =   3840
         Width           =   7575
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   3240
            TabIndex        =   59
            Top             =   1080
            Width           =   4215
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   26
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   1200
            TabIndex        =   61
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   24
            Left            =   120
            TabIndex        =   62
            Tag             =   "Código do Contato"
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "frmFO.frx":48E4
            TabIndex        =   121
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":494E
            TabIndex        =   120
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmFO.frx":49BE
            TabIndex        =   119
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":4A26
            TabIndex        =   118
            Top             =   240
            Width           =   615
         End
         Begin ZEUS.chameleonButton chameleonButton9 
            Height          =   255
            Left            =   7080
            TabIndex        =   58
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BTYPE           =   4
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
            MICON           =   "frmFO.frx":4A92
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
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   2
         Left            =   -74280
         TabIndex        =   47
         Tag             =   "Editar registro"
         ToolTipText     =   "Editar registro"
         Top             =   3900
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
         MICON           =   "frmFO.frx":4AAE
         PICN            =   "frmFO.frx":4ACA
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
         Left            =   -74880
         TabIndex        =   9
         Tag             =   "Inserir registro"
         ToolTipText     =   "Inserir registro"
         Top             =   3900
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
         MICON           =   "frmFO.frx":57A4
         PICN            =   "frmFO.frx":57C0
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
         Caption         =   "Status"
         Height          =   615
         Left            =   -72840
         TabIndex        =   48
         Top             =   3960
         Width           =   1695
         Begin VB.Label Label37 
            Caption         =   "cont"
            Height          =   255
            Left            =   1200
            TabIndex        =   50
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label36 
            Caption         =   "Inclusão"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   3
         Left            =   -63000
         TabIndex        =   46
         Tag             =   "Gerar resumo"
         ToolTipText     =   "Gerar resumo"
         Top             =   3960
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
         MICON           =   "frmFO.frx":649A
         PICN            =   "frmFO.frx":64B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame17 
         Caption         =   "Totais "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -64560
         TabIndex        =   45
         Top             =   1920
         Width           =   2175
         Begin ACTIVESKINLibCtl.SkinLabel lblTotPint 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":7190
            TabIndex        =   105
            Top             =   960
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel lblTotal 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":71EE
            TabIndex        =   104
            Top             =   480
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":724C
            TabIndex        =   103
            Top             =   720
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":72CA
            TabIndex        =   102
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Informações "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74760
         TabIndex        =   37
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   39
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   40
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox cbocadastro 
            Height          =   315
            Left            =   1560
            TabIndex        =   40
            Text            =   "KG"
            Top             =   1320
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   2520
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   1920
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   503
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":733E
            TabIndex        =   134
            Top             =   2280
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":73B4
            TabIndex        =   133
            Top             =   1680
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "frmFO.frx":7430
            TabIndex        =   132
            Top             =   1080
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":749E
            TabIndex        =   131
            Top             =   1080
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":7512
            TabIndex        =   130
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   -67440
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   33
         Top             =   3240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
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
      Begin VB.Frame Frame10 
         Caption         =   "Dados do resumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   7095
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   32
            Left            =   1200
            TabIndex        =   24
            Top             =   480
            Width           =   5055
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   31
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmFO.frx":7588
            TabIndex        =   123
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":75FA
            TabIndex        =   122
            Top             =   240
            Width           =   735
         End
         Begin VB.Frame Frame11 
            Caption         =   "Cálculo por"
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
            TabIndex        =   26
            Top             =   840
            Width           =   6855
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   37
               Left            =   120
               TabIndex        =   34
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtcadastro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
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
               Height          =   285
               Index           =   38
               Left            =   1200
               TabIndex        =   36
               Top             =   1200
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
               Height          =   255
               Left            =   1200
               OleObjectBlob   =   "frmFO.frx":7666
               TabIndex        =   125
               Top             =   960
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":76CE
               TabIndex        =   124
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   36
               Left            =   6240
               TabIndex        =   32
               Top             =   600
               Width           =   375
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Alíquota"
               Height          =   210
               Left            =   5160
               TabIndex        =   31
               Top             =   690
               Width           =   975
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   35
               Left            =   2640
               TabIndex        =   30
               Top             =   600
               Width           =   2415
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   34
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   2415
            End
            Begin VB.OptionButton optCadastro 
               Caption         =   "Peso:"
               Height          =   375
               Index           =   3
               Left            =   2640
               TabIndex        =   28
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optCadastro 
               Caption         =   "Dimensão:"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   33
            Left            =   6360
            TabIndex        =   25
            Top             =   480
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   11
         Top             =   4680
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   5106
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483646
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
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000007&
         Height          =   1185
         Left            =   -64440
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   2010
      End
      Begin VB.Frame Frame8 
         Caption         =   "Dados do material "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   0
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   10215
         Begin VB.Frame Frame4 
            Caption         =   "Desenho/revisão "
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
            TabIndex        =   83
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   29
               Left            =   4320
               TabIndex        =   86
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   84
               Top             =   480
               Width           =   4095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "frmFO.frx":7742
               TabIndex        =   96
               Top             =   240
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":77AA
               TabIndex        =   95
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   3240
               TabIndex        =   85
               Top             =   120
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de material "
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
            Left            =   5160
            TabIndex        =   79
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   10
               Left            =   1200
               TabIndex        =   81
               Top             =   480
               Width           =   3135
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   80
               Tag             =   "Código do tipo de material"
               Top             =   480
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   1200
               OleObjectBlob   =   "frmFO.frx":7818
               TabIndex        =   98
               Top             =   240
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":788A
               TabIndex        =   97
               Top             =   240
               Width           =   615
            End
            Begin ZEUS.chameleonButton chameleonButton7 
               Height          =   255
               Left            =   4440
               TabIndex        =   82
               Top             =   480
               Width           =   375
               _extentx        =   661
               _extenty        =   450
               btype           =   4
               tx              =   "..."
               enab            =   -1  'True
               font            =   "frmFO.frx":78F6
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   13160660
               bcolo           =   13160660
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frmFO.frx":7922
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   0
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Parâmetro/Cálculo Área de pintura"
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
            Left            =   6840
            TabIndex        =   44
            Top             =   2400
            Width           =   3255
            Begin VB.CheckBox Check1 
               Caption         =   "Pintura:"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   5
               Left            =   120
               TabIndex        =   77
               Tag             =   "Pintura"
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Material"
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
            TabIndex        =   43
            Top             =   1320
            Width           =   9975
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   4
               Left            =   8160
               TabIndex        =   4
               Tag             =   "Quant. CJ"
               ToolTipText     =   "Quant. CJ"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   960
               TabIndex        =   1
               Top             =   480
               Width           =   7095
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   0
               Tag             =   "Código do Material"
               Top             =   480
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Left            =   8160
               OleObjectBlob   =   "frmFO.frx":7940
               TabIndex        =   101
               Top             =   240
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
               Height          =   255
               Left            =   960
               OleObjectBlob   =   "frmFO.frx":79B0
               TabIndex        =   100
               Top             =   240
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":7A22
               TabIndex        =   99
               Top             =   240
               Width           =   615
            End
            Begin ZEUS.chameleonButton chameleonButton5 
               Height          =   255
               Left            =   9480
               TabIndex        =   3
               Top             =   480
               Width           =   375
               _extentx        =   661
               _extenty        =   450
               btype           =   4
               tx              =   "..."
               enab            =   -1  'True
               font            =   "frmFO.frx":7A8E
               coltype         =   1
               focusr          =   -1  'True
               bcol            =   13160660
               bcolo           =   13160660
               fcol            =   0
               fcolo           =   0
               mcol            =   12632256
               mptr            =   1
               micon           =   "frmFO.frx":7ABA
               umcol           =   -1  'True
               soft            =   0   'False
               picpos          =   0
               ngrey           =   0   'False
               fx              =   0
               hand            =   0   'False
               check           =   0   'False
               value           =   0   'False
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   9000
               TabIndex        =   2
               Tag             =   "Quant. CJ"
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Cálculo por "
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
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   2400
            Width           =   6615
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   8
               Left            =   3360
               TabIndex        =   8
               Tag             =   "Peso"
               Top             =   480
               Width           =   3015
            End
            Begin VB.OptionButton optCadastro 
               Caption         =   "Peso:"
               Height          =   255
               Index           =   1
               Left            =   3360
               TabIndex        =   7
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   6
               Tag             =   "Dimensão"
               Top             =   480
               Width           =   3135
            End
            Begin VB.OptionButton optCadastro 
               Caption         =   "Dimensão:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Label Label19 
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados da FO "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7155
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   4815
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   2880
            TabIndex        =   52
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   94633985
            CurrentDate     =   40449
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   1320
            TabIndex        =   53
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   94633985
            CurrentDate     =   40449
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   55
            Tag             =   "Nº Ficha de Orçamento"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   30
            Left            =   120
            TabIndex        =   54
            Tag             =   "Descrição do serviço"
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   56
            Tag             =   "Nº da SDC"
            ToolTipText     =   "Nº da Solicitação de Cotação"
            Top             =   1680
            Width           =   4575
         End
         Begin VB.TextBox txtcadastro 
            Height          =   3615
            Index           =   28
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   76
            Tag             =   "Observação"
            Top             =   3360
            Width           =   4575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":7AD8
            TabIndex        =   94
            Top             =   3120
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":7B4A
            TabIndex        =   91
            Top             =   1440
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":7BB6
            TabIndex        =   90
            Top             =   840
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmFO.frx":7C28
            TabIndex        =   89
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmFO.frx":7C96
            TabIndex        =   88
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":7D06
            TabIndex        =   87
            Top             =   240
            Width           =   615
         End
         Begin VB.Frame Frame5 
            Caption         =   "FCE - Ficha de Controle de Encomenda "
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
            TabIndex        =   14
            Top             =   2040
            Width           =   4575
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":7D70
               TabIndex        =   135
               Top             =   480
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel Label32 
               Height          =   255
               Left            =   1920
               OleObjectBlob   =   "frmFO.frx":7DCC
               TabIndex        =   93
               Top             =   480
               Width           =   2415
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":7E2A
               TabIndex        =   92
               Top             =   240
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   1920
               OleObjectBlob   =   "frmFO.frx":7E96
               TabIndex        =   51
               Top             =   240
               Width           =   855
            End
         End
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   -71040
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         Height          =   2535
         Left            =   -67560
         Top             =   480
         Width           =   5160
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000C0&
         FillColor       =   &H80000006&
         Height          =   1455
         Left            =   -64560
         Top             =   420
         Width           =   2160
      End
   End
End
Attribute VB_Name = "frmFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsLocal As New ADODB.Recordset
Private rsMaterial As New ADODB.Recordset
Private const0(9) As Double
Private vAr0(9) As Double
Private X As Integer, Y As Integer
Private Conta As Integer
Private Formula As String
Private ForPint As String
Private SqlM As String
Private SomaTotal As Double
Private SomaPint As Double
Private QuantCJ As Double
Private PesoTotal As Double
Private TipoCad As String

Private Sub chamCad_Click(Index As Integer)
    Select Case Index
    Case 0
        txtCadastro_KeyDown 2, 13, 1
        LimpaControles
        SomaListview
        txtCadastro(0).SetFocus
    Case 1
        ExcluirItem
        SomaListview
        txtCadastro(0).SetFocus
    Case 2
        AlterarItem1
    Case 3
        GerarResumo
        SSTab1.Tab = 2
        optCadastro(3).Value = True
        Check3.Value = 1
        Msgbox "Resumo gerado com sucesso"
    Case 4
        GravarDados
        txtCadastro(0).SetFocus
    Case 5
        ExportaExcel
        Msgbox "Dados exportados com sucesso", vbInformation, "Zeus"
    Case 6
        If Msgbox("Deseja sair da tela de cadastro?", vbQuestion + vbYesNo, "Zeus") = vbYes Then
            'CancelaSN = 1
            Unload Me
        End If
    End Select
End Sub

Private Sub chameleonButton5_Click()
    ChamaGrid
    CarregaDados (0)
End Sub

Private Sub chameleonButton7_Click()
    ChamaGridMat
    CarregaTipoMat
End Sub

Private Sub chameleonButton8_Click()
    txtCadastro(13) = ""
    ChamaGridCli
End Sub

Private Sub chameleonButton9_Click()
    txtCadastro(24) = ""
    ChamaGridCont
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtCadastro(5).Enabled = True
    Else
        txtCadastro(5).Enabled = False
    End If
End Sub

'Private Sub Check2_Click()
'    If Check2.Value = 0 Then
'        Option1.Value = False
'        Option2.Value = False
'        Option3.Value = False
'        Option1.Enabled = False
'        Option2.Enabled = False
'        Option3.Enabled = False
'        SkinLabel45 = ""
'        SkinLabel45.Enabled = False
'    Else
'        Option1.Enabled = True
'        Option2.Enabled = True
'        Option3.Enabled = True
'        SkinLabel45.Enabled = True
'    End If
'End Sub

Private Sub Check3_Click()
    If Check3.Value = 0 Then
        txtCadastro(36).Enabled = False
        txtCadastro(36).BackColor = &H80000004
    Else
        txtCadastro(36).Enabled = True
        txtCadastro(36).BackColor = &H80000005
    End If
End Sub

Private Sub chamCad_MouseOver(Index As Integer)
    Legenda = chamCad(Index).ToolTipText
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub chamCad_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
'    frmFO.Top = MDIPrincipal.pctOrc.Height + (MDIPrincipal.pctOrc.Height * 50 / 100)
'    frmFO.Left = 110
    DTPicker1 = Date
    DTPicker2 = Date
    
    SSTab1.Tab = 0
    SomaTotal = 0
    SomaPint = 0
    listview_cabecalho
    TipoCad = Pesquisa
    If TipoCad = "novo" Then
        LimpaControles
        txtCadastro(6) = Format(GeraCodigo, "000000") & ""
        txtCadastro(6).Enabled = False
        optCadastro_Click (0)
    ElseIf TipoCad = "editar" Then
        ResultPesq
        DesbloqueiaControles
        txtCadastro_KeyDown 6, 13, 6
    End If
    optCadastro_Click 1
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub
Private Sub LimpaControles()
    Formula = ""
    ForPint = ""
    Conta = 0
End Sub

Private Sub LimpaControles1()
    Dim X As Integer
    For X = 0 To 5
        txtCadastro(X) = ""
    Next
    For X = 7 To 38
        txtCadastro(X) = ""
    Next
'    Option1.Value = False
'    Option2.Value = False
'    Option3.Value = False
'    Check2.Value = 0
    txtCadastro(8) = ""
    ListView1.ListItems.Clear
'    DTPicker1 = ""
    Formula = ""
    ForPint = ""
    Conta = 0
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Material", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Dimensões", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Peso Unit/Qtd.", ListView1.Width / 7.6
    ListView1.ColumnHeaders.Add , , "Un", ListView1.Width / 32
    ListView1.ColumnHeaders.Add , , "Q.CJ", ListView1.Width / 19.5
    ListView1.ColumnHeaders.Add , , "Peso Total", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Área Pint.", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Desenho", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "codigo+material", ListView1.Width / 1000
    ListView1.ColumnHeaders.Add , , "Rev", ListView1.Width / 23
    ListView1.ColumnHeaders.Add , , "sequencia", ListView1.Width / 1000
    
    ListView2.ColumnHeaders.Add , , "Item", ListView2.Width / 16
    ListView2.ColumnHeaders.Add , , "Código", ListView2.Width / 16
    ListView2.ColumnHeaders.Add , , "Descrição", ListView2.Width / 5
    ListView2.ColumnHeaders.Add , , "Material", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Un", ListView2.Width / 32
    ListView2.ColumnHeaders.Add , , "Peso Unit/Qtd.", ListView2.Width / 7.6
    ListView2.ColumnHeaders.Add , , "Área Pint.", ListView2.Width / 10
    
    ListView2.ColumnHeaders.Add , , "Dimensões resumo", ListView2.Width / 6.8
    ListView2.ColumnHeaders.Add , , "Peso MP", ListView2.Width / 12
    
    ListView2.ColumnHeaders.Add , , "Alíq.", ListView2.Width / 12
    ListView2.ColumnHeaders.Add , , "Qtd.", ListView2.Width / 12
    ListView2.ColumnHeaders.Add , , "Status", ListView2.Width / 12
    
    
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    
    Me.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    
    ListView1.View = lvwReport
    ListView2.View = lvwReport
End Sub
Private Sub IncluirItem()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, ProxSeq As Integer
    If ValidaCampo2 = False Then Exit Sub
    
    'Variavel do sistema para calculo da area de pintura, consta na formula de pintura
    If optCadastro(0).Value = True Then
        PesoTotal = Format(ScriptControl1.Eval(Formula) * Me.txtCadastro(4), "#,##0.000;(#,##0.000)")
    Else
        PesoTotal = Format(txtCadastro(8) * txtCadastro(4), "#,##0.000;(#,##0.000)")
    End If
    Y = ListView1.ListItems.Count
    
    If Label36.Caption = "Alteração" Then
        ListView1.ListItems(Val(Label37)).Selected = True
        ListView1.ListItems(Val(Label37)).EnsureVisible
    
        'ListView1.SelectedItem.ListSubItems.Item (1)
        Label36.Caption = "Inclusão"
        If Check1.Value = 1 Then
            'Variavel q contem a formula para calcular a área de pintura
            'O Replace esta sendo aplicado aki pq so agora q foi encontrado o PesoTotal
            ForPint = Replace(ForPint, "pesototal", PesoTotal)
            ForPint = Replace(ForPint, ",", ".")
            If Me.txtCadastro(2) <> ListView1.SelectedItem.ListSubItems.Item(3) Or Me.txtCadastro(8) <> ListView1.SelectedItem.ListSubItems.Item(3) Then
                ListView1.SelectedItem.ListSubItems.Item(8) = Format(ScriptControl1.Eval(ForPint) * txtCadastro(5), "#,##0.000;(#,##0.000)")
            End If
        End If
        ListView1.SelectedItem.ListSubItems.Item(1) = Me.txtCadastro(1).Text
        ListView1.SelectedItem.ListSubItems.Item(2) = Me.txtCadastro(9).Text & "-" & Me.txtCadastro(10).Text
        
        If Me.txtCadastro(2) <> ListView1.SelectedItem.ListSubItems.Item(3) Then
'        If Me.txtcadastro(2) <> ListView1.SelectedItem.ListSubItems.Item(3) Or Me.txtcadastro(8) <> ListView1.SelectedItem.ListSubItems.Item(4) Then
            If optCadastro(0).Value = True Then
                ListView1.SelectedItem.ListSubItems.Item(4) = Format(ScriptControl1.Eval(Formula), "#,##0.000;(#,##0.000)")
                ListView1.SelectedItem.ListSubItems.Item(7) = Format(ScriptControl1.Eval(Formula) * Me.txtCadastro(4), "#,##0.000;(#,##0.000)")
            Else
                ListView1.SelectedItem.ListSubItems.Item(4) = Format(txtCadastro(8), "#,##0.000;(#,##0.000)")
                ListView1.SelectedItem.ListSubItems.Item(7) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
            End If
        End If
        
        If Me.txtCadastro(8) <> ListView1.SelectedItem.ListSubItems.Item(4) Then
            If optCadastro(1).Value = True Then
                ListView1.SelectedItem.ListSubItems.Item(4) = Format(txtCadastro(8), "#,##0.000;(#,##0.000)")
                ListView1.SelectedItem.ListSubItems.Item(7) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
            End If
        End If
        
        ListView1.SelectedItem.ListSubItems.Item(3) = Me.txtCadastro(2).Text
        ListView1.SelectedItem.ListSubItems.Item(5) = Me.txtCadastro(3).Text
        ListView1.SelectedItem.ListSubItems.Item(6) = Me.txtCadastro(4).Text
        ListView1.SelectedItem.ListSubItems.Item(9) = Me.txtCadastro(7).Text
        ListView1.SelectedItem.ListSubItems.Item(12) = Format(Label37, "0000") 'Me.txtcadastro(1).Text & Me.txtcadastro(10).Text
        ListView1.SelectedItem.ListSubItems.Item(11) = Me.txtCadastro(29).Text
    Else
        'Ordena Listview pela sequencia de cadastramento antes de gravar
        Me.ListView1.Sorted = True
        Me.ListView1.SortKey = 12
        Me.ListView1.SortOrder = lvwAscending
        '------
        If ListView1.ListItems.Count > 0 Then
            ListView1.ListItems(ListView1.ListItems.Count).Selected = True
            ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
            ProxSeq = Val(ListView1.SelectedItem.ListSubItems.Item(12)) + 1
        Else
            ProxSeq = 1
        End If
        
        Set ItemLst = ListView1.ListItems.Add(, , Format(txtCadastro(0), "000000"))
        Label36.Caption = "Inclusão"
        If Check1.Value = 1 Then
            'Variavel q contem a formula para calcular a área de pintura
            'O Replace esta sendo aplicado aki pq so agora q foi encontrado o PesoTotal
            ForPint = Replace(ForPint, "pesototal", PesoTotal)
            ForPint = Replace(ForPint, ",", ".")
            ItemLst.SubItems(8) = Format(ScriptControl1.Eval(ForPint) * txtCadastro(5), "#,##0.000;(#,##0.000)")
        End If
        ItemLst.SubItems(1) = Me.txtCadastro(1).Text
        ItemLst.SubItems(2) = Me.txtCadastro(9).Text & "-" & Me.txtCadastro(10).Text
        ItemLst.SubItems(3) = Me.txtCadastro(2).Text
        If optCadastro(0).Value = True Then
            ItemLst.SubItems(4) = Format(ScriptControl1.Eval(Formula), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(7) = Format(ScriptControl1.Eval(Formula) * Me.txtCadastro(4), "#,##0.000;(#,##0.000)")
        Else
            ItemLst.SubItems(4) = Format(txtCadastro(8), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(7) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
        End If
        ItemLst.SubItems(5) = Me.txtCadastro(3).Text
        ItemLst.SubItems(6) = Me.txtCadastro(4).Text
        ItemLst.SubItems(9) = Me.txtCadastro(7).Text
        ItemLst.SubItems(10) = Me.txtCadastro(1).Text & Me.txtCadastro(10).Text
        ItemLst.SubItems(11) = Me.txtCadastro(29).Text
        ItemLst.SubItems(12) = Format(ProxSeq, "0000")
        
        ListView1.ListItems(ListView1.ListItems.Count).Selected = True
        ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
        
    End If
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    If txtCadastro(3) <> "pç" And txtCadastro(3) <> "PÇ" Then
        If optCadastro(0).Value = True Then
            SomaTotal = SomaTotal + ScriptControl1.Eval(Formula) * Me.txtCadastro(4)
        Else
            SomaTotal = Format(SomaTotal + txtCadastro(8) * Me.txtCadastro(4), "#,##0.000;(#,##0.000")
        End If
        
        If Check1.Value = 1 Then SomaPint = SomaPint + ScriptControl1.Eval(ForPint) * Me.txtCadastro(5)
    End If
    'lblTotal.Caption = Format(SomaTotal, "#,##0.0;(#,##0.0)") 'Format(SomaTotal, "#,##0.000000000;(#,##0.000000000)")
    'lblTotPint.Caption = Format(SomaPint, "#,##0.00;(#,##0.00)")
    txtCadastro(0) = ""
    txtCadastro(1) = ""
    txtCadastro(2) = ""
    Text1.Text = ""
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, verifique se as dimensões digitadas estão de acordo com as referidas na fórmula!", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub IncluirItem2()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaCampo = False Then Exit Sub
    
    'Variavel do sistema para calculo da area de pintura, consta na formula de pintura
    If optCadastro(2).Value = True Then
        PesoTotal = Format(ScriptControl1.Eval(Formula) * Me.txtCadastro(37), "#,##0.000;(#,##0.000)")
    Else
        PesoTotal = Format(txtCadastro(35), "#,##0.000;(#,##0.000)")
    End If
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView2.ListItems.Item(X) = Me.txtCadastro(38) Then
                If optCadastro(2).Value = True Then
                    ListView2.SelectedItem.ListSubItems.Item(7) = Me.txtCadastro(34).Text
                    ListView2.SelectedItem.ListSubItems.Item(8) = Format(ScriptControl1.Eval(Formula) * Me.txtCadastro(37), "#,##0.000;(#,##0.000)")
                    ListView2.SelectedItem.ListSubItems.Item(9) = 0
                    ListView2.SelectedItem.ListSubItems.Item(10) = txtCadastro(37)
                Else
                    ListView2.SelectedItem.ListSubItems.Item(7) = "-"
                    If Check3.Value = 1 Then ListView2.SelectedItem.ListSubItems.Item(8) = Format((txtCadastro(35) * txtCadastro(36) / 100) + txtCadastro(35), "#,##0.000;(#,##0.000)")
                    If Check3.Value = 0 Then ListView2.SelectedItem.ListSubItems.Item(8) = Format(txtCadastro(35), "#,##0.000;(#,##0.000)")
                    If Check3.Value = 1 Then ListView2.SelectedItem.ListSubItems.Item(9) = txtCadastro(36)
                    ListView2.SelectedItem.ListSubItems.Item(10) = 1
                End If
                ListView2.SelectedItem.ListSubItems.Item(11) = "Alterado"
                
                Exit For
            End If
        Next
    End If
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    Text4.Text = ""
    For X = 31 To 38
        txtCadastro(X) = ""
    Next
    LimpaControles
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, verifique se as dimensões digitadas estão de acordo com as referidas na fórmula!", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub ExcluirItem()
On Error GoTo TrataErro
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    'If ListView1.SelectedItem.ListSubItems.Item(4) <> "pç" Then
    '    SomaTotal = SomaTotal - ListView1.SelectedItem.ListSubItems.Item(6)
    '    SomaPint = SomaPint - ListView1.SelectedItem.ListSubItems.Item(7)
    '    lblTotal.Caption = Format(SomaTotal, "#,##0.0;(#,##0.0)") 'Format(SomaTotal, "#,##0.000000000;(#,##0.000000000)")
    '    lblTotPint.Caption = Format(SomaPint, "#,##0.00;(#,##0.00)")
    'End If
    ListView1.ListItems.Remove (X)
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, Selecione um item antes de excluir", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub AlterarItem()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadastro(31).Text = ListView2.SelectedItem.ListSubItems.Item(1)
    Me.txtCadastro(32).Text = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.txtCadastro(33).Text = ListView2.SelectedItem.ListSubItems.Item(4)
    If ListView2.SelectedItem.ListSubItems.Item(8) = 0 Then Me.txtCadastro(35).Text = ListView2.SelectedItem.ListSubItems.Item(5) Else Me.txtCadastro(35).Text = ListView2.SelectedItem.ListSubItems.Item(8)
    If ListView2.SelectedItem.ListSubItems.Item(9) = 0 Then Me.txtCadastro(36).Text = 5 Else Me.txtCadastro(36).Text = ListView2.SelectedItem.ListSubItems.Item(9)
    Me.txtCadastro(37).Text = ListView2.SelectedItem.ListSubItems.Item(10)
    Me.txtCadastro(38).Text = ListView2.ListItems.Item(X)
    CarregaDados (31)
    txtCadastro(31).Enabled = False
    txtCadastro(32).Enabled = False
    txtCadastro(33).Enabled = False
End Sub

Private Sub AlterarItem1()
    Label36.Caption = "Alteração"
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Label37 = Val(ListView1.SelectedItem.ListSubItems.Item(12))
    Me.txtCadastro(7) = ListView1.SelectedItem.ListSubItems.Item(9)
    Me.txtCadastro(29) = ListView1.SelectedItem.ListSubItems.Item(11)
    Me.txtCadastro(9) = Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 6)
    Me.txtCadastro(0) = ListView1.ListItems.Item(X)
    Me.txtCadastro(1) = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.txtCadastro(3) = ListView1.SelectedItem.ListSubItems.Item(5)
    Me.txtCadastro(2) = ListView1.SelectedItem.ListSubItems.Item(3)
    Me.txtCadastro(4) = ListView1.SelectedItem.ListSubItems.Item(6)
    
    If ListView1.SelectedItem.ListSubItems.Item(3) = "-" Then
        Me.optCadastro(1).Value = True
        optCadastro_Click (1)
        Me.txtCadastro(8) = ListView1.SelectedItem.ListSubItems.Item(4)
        txtCadastro(8).SetFocus
        Check1.Value = 0
        txtCadastro(5).Text = ""
        If Check1.Value = 0 Then txtCadastro(5).Enabled = False
        txtCadastro(0).BackColor = &HC0C0FF
    End If
    If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then
        Me.optCadastro(0).Value = True
        optCadastro_Click (0)
        txtCadastro(2).SetFocus
        txtCadastro(0).BackColor = &H80000005
    End If
    
    If Val(txtCadastro(0)) = 0 Then
        optCadastro_Click (1)
    End If
    
    If ListView1.SelectedItem.ListSubItems.Item(8) = "0,00" Then Check1.Value = 0 Else Check1.Value = 1
    txtCadastro_KeyDown 0, 13, 0
    txtCadastro_KeyDown 9, 13, 9
    
    If txtCadastro(0) = "000000" Then
        txtCadastro(8).BackColor = &HC0C0FF
        Me.txtCadastro(1) = ListView1.SelectedItem.ListSubItems.Item(1)
    End If

End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsDeleta As New ADODB.Recordset
    Dim rsGravaLM As New ADODB.Recordset
    Dim rsGravaFO As New ADODB.Recordset
    Dim rsGravaResumo As New ADODB.Recordset
    
    Dim sqlExc As String
    Dim sql As String
    Dim Y As Integer, X As Integer
    cnBanco.BeginTrans

    sql = "Select * from tbListaMaterial order by codfo"
    rsGravaLM.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
    sqlExc = "Delete from tbListaMaterial where tbListaMaterial.codfo = '" & Val(txtCadastro(6)) & "'"
    rsDeleta.Open sqlExc, cnBanco
    
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaLM.AddNew
        rsGravaLM(0) = txtCadastro(6)
        rsGravaLM(1) = ListView1.SelectedItem.ListSubItems.Item(12)
        rsGravaLM(2) = ListView1.SelectedItem.ListSubItems.Item(9)
        
        rsGravaLM(3) = Val(ListView1.ListItems.Item(X))
        rsGravaLM(4) = ListView1.SelectedItem.ListSubItems.Item(6)
        rsGravaLM(5) = ListView1.SelectedItem.ListSubItems.Item(3)
        rsGravaLM(6) = ListView1.SelectedItem.ListSubItems.Item(4)
        If ListView1.SelectedItem.ListSubItems.Item(8) <> "" Then rsGravaLM(7) = ListView1.SelectedItem.ListSubItems.Item(8) Else rsGravaLM(7) = 0
        rsGravaLM(8) = Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 6))
        rsGravaLM(9) = ListView1.SelectedItem.ListSubItems.Item(11)
        If Val(ListView1.ListItems.Item(X)) = 0 Then rsGravaLM(10) = ListView1.SelectedItem.ListSubItems.Item(1)
    Next
    If Not rsGravaLM.EOF Then rsGravaLM.Update
    
    sql = "Select * from tbFo where tbfo.codfo = '" & Val(txtCadastro(6)) & "'"
    rsGravaFO.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If rsGravaFO.EOF Then
        rsGravaFO.AddNew
    End If
    rsGravaFO(0) = txtCadastro(6)
    rsGravaFO(1) = Format(DTPicker1, "dd/mm/yyyy")

    If rsGravaFO(2) = "" Then rsGravaFO(2) = 1
    If rsGravaFO(2) = 1 Then rsGravaFO(2) = 1
    If rsGravaFO(2) = 2 Then rsGravaFO(2) = 2
    If rsGravaFO(2) = 3 Then rsGravaFO(2) = 3
    
    rsGravaFO(4) = txtCadastro(11)
    If txtCadastro(13) <> "" Then rsGravaFO(5) = txtCadastro(13)
    If txtCadastro(24) <> "" Then rsGravaFO(6) = txtCadastro(24)
    rsGravaFO(7) = txtCadastro(28)
    rsGravaFO(8) = txtCadastro(30)
    If DTPicker2 <> "" Then
        rsGravaFO(9) = Format(DTPicker2, "dd/mm/yyyy")
    End If
    rsGravaFO(10) = txtCadastro(39)
    If txtCadastro(40) <> "" Then rsGravaFO(11) = Format(txtCadastro(40), "#,##0.000;(#,##0.000)")
    rsGravaFO(12) = cboCadastro
    mskCadastro(1).PromptInclude = False
    If mskCadastro(1) <> "" Then rsGravaFO(13) = mskCadastro(1)
    mskCadastro(1).PromptInclude = True
    rsGravaFO(14) = "S"
    
    If Not rsGravaFO.EOF Then rsGravaFO.Update

'***************** INICIO GRAVAR DADOS DA TABELA DE RESUMO ***************
'Grava apenas se houver alguma informação na tabela de resumo
    If ListView2.ListItems.Count <> 0 Then
        
        sqlExc = "Delete from tbResumo where tbResumo.codfo = '" & Val(txtCadastro(6)) & "'"
        rsDeleta.Open sqlExc, cnBanco
        
        sql = "Select * from tbResumo where tbResumo.codfo = '" & Val(txtCadastro(6)) & "'"
        rsResumo.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
        
        Y = ListView2.ListItems.Count
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
            
            
            rsResumo.AddNew
            rsResumo.Fields(0) = txtCadastro(6)
            rsResumo.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
            rsResumo.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(7)
            rsResumo.Fields(3) = ListView2.SelectedItem.ListSubItems.Item(8)
            rsResumo.Fields(4) = ListView2.SelectedItem.ListSubItems.Item(9)
            rsResumo.Fields(5) = ListView2.SelectedItem.ListSubItems.Item(10)
            rsResumo.Fields(6) = ListView2.SelectedItem.ListSubItems.Item(5)
            If ListView2.SelectedItem.ListSubItems.Item(4) <> "-" Then rsResumo.Fields(7) = Val(Mid$(ListView2.SelectedItem.ListSubItems.Item(3), 1, 6)) Else rsResumo.Fields(7) = 0
            rsResumo.Fields(8) = X
'            If Val(ListView2.SelectedItem.ListSubItems.Item(1)) = 0 Then rsResumo.Fields(8) = X Else rsResumo.Fields(8) = Val(ListView2.SelectedItem.ListSubItems.Item(1))
            If Val(ListView2.SelectedItem.ListSubItems.Item(1)) = 0 Then rsResumo.Fields(9) = ListView2.SelectedItem.ListSubItems.Item(2)
        
        Next
        If Not rsResumo.EOF Then rsResumo.Update
        rsResumo.Close
    End If

'***************** FIM GRAVAR DADOS DA TABELA DE RESUMO ******************
    
    cnBanco.CommitTrans
    rsGravaFO.Close
    rsGravaLM.Close
AtualizaListview
    Msgbox "Dados gravados com sucesso", vbInformation, "Zeus"
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
    AlterarItem1
End Sub

Private Sub ListView2_DblClick()
    If txtCadastro(34).Enabled = True Then txtCadastro(34).SetFocus
    If txtCadastro(35).Enabled = True Then txtCadastro(35).SetFocus
    AlterarItem
End Sub

Private Sub mskCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtCadastro(7).SetFocus
        End If
    End If
End Sub

Private Sub optCadastro_Click(Index As Integer)
    If optCadastro(0).Value = True Then
        txtCadastro(2).Enabled = True
        txtCadastro(8).Enabled = False
        txtCadastro(8).BackColor = &H80000004
        txtCadastro(2).BackColor = &H80000005
    End If
    If optCadastro(1).Value = True Then
        txtCadastro(8).Enabled = True
        txtCadastro(2).Enabled = False
        txtCadastro(2).BackColor = &H80000004
        txtCadastro(8).BackColor = &H80000005
    End If

    If optCadastro(2).Value = True Then
        txtCadastro(34).Enabled = True
        txtCadastro(35).Enabled = False
        txtCadastro(36).Enabled = False
        txtCadastro(37).Enabled = True
        txtCadastro(36).BackColor = &H80000004
        txtCadastro(35).BackColor = &H80000004
        txtCadastro(34).BackColor = &H80000005
        txtCadastro(37).BackColor = &H80000005
        Check3.Value = 0
        Check3.Enabled = False
    End If
    If optCadastro(3).Value = True Then
        txtCadastro(34).Enabled = False
        txtCadastro(35).Enabled = True
        txtCadastro(37).Enabled = False
        txtCadastro(35).BackColor = &H80000005
        txtCadastro(34).BackColor = &H80000004
        txtCadastro(37).BackColor = &H80000004
        Check3.Enabled = True
        Check3.Value = 0
    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value = False Then
        Frame14.Enabled = False
        Label32.Enabled = False
        Label33.Enabled = False
        SkinLabel45.Enabled = False
        txtCadastro(41).Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option3.Value = False Then
        Frame14.Enabled = False
        Label32.Enabled = False
        Label33.Enabled = False
        SkinLabel45.Enabled = False
        txtCadastro(41).Enabled = False
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        Frame14.Enabled = True
        Label32.Enabled = True
        Label33.Enabled = True
        SkinLabel45.Enabled = True
        txtCadastro(41).Enabled = True
    End If
'    SkinLabel45.SetFocus
End Sub

Private Sub txtCadastro_GotFocus(Index As Integer)
    If Index = 4 Then
        txtCadastro(4).SelStart = 0
        txtCadastro(4).SelLength = Len(txtCadastro(4).Text)
    End If
    Dim X As Integer
    For X = 1 To 11
        txtCadastro(X).SelStart = 0
        txtCadastro(X).SelLength = Len(txtCadastro(X).Text)
    Next
    For X = 13 To 40
        txtCadastro(X).SelStart = 0
        txtCadastro(X).SelLength = Len(txtCadastro(X).Text)
    Next
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCadastro(0) = "" Then ChamaGrid
            CarregaDados (Index)
        End If
    ElseIf Index = 1 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtCadastro(4).SetFocus
        End If
    ElseIf Index = 2 Or Index = 34 Or Index = 35 Then
        If KeyCode = &H8 Then
            txtCadastro(2) = ""
            Formula = ""
            ForPint = ""
            Conta = 0
            CarregaDados (0)
            txtCadastro(2).SetFocus
        End If
        If Conta > 0 Then
            If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
                If Index = 2 Then CapVar
                If Index = 34 Then CapVar2
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
                
                QuantCJ = Val(txtCadastro(4))
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
                If Index = 2 Then IncluirItem
                If Index = 34 Or Index = 35 Then IncluirItem2
                Conta = 0
                LimpaControles
                If Index = 2 Then txtCadastro(0).SetFocus
            End If
            If KeyCode = &H6D Then
                If Index = 2 Then CapVar
                If Index = 34 Or Index = 35 Then CapVar2
            End If
        ElseIf Conta = 0 Then
            If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
                vAr0(1) = Val(txtCadastro(Index))
                Text2.Text = vAr0(1)
                Y = X
                X = Len(txtCadastro(Index))
                Conta = Conta + 1
                Text2.Text = Formula
                txtCadastro_KeyDown Index, 13, 1
            End If
            If KeyCode = &H6D Then 'traço
                vAr0(1) = Val(txtCadastro(Index))
                Text2.Text = vAr0(1)
                Y = X
                X = Len(txtCadastro(Index))
                Conta = Conta + 1
            End If
        End If
        SomaListview
    ElseIf Index = 4 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If optCadastro(0).Value = True Then
                If Val(txtCadastro(0)) <> 0 Then txtCadastro(2).SetFocus
            Else
                txtCadastro(8).SetFocus
            End If
        End If
    ElseIf Index = 6 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCadastro(6) = "" Or Val(txtCadastro(6)) = 0 Then Exit Sub
            CarregaFO
            txtCadastro(6) = Format(txtCadastro(6), "000000")
            If txtCadastro(13) <> "" Then txtCadastro_KeyDown 13, 13, 13
            If txtCadastro(24) <> "" Then txtCadastro_KeyDown 24, 13, 24
        End If
    ElseIf Index = 7 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtCadastro(0).SetFocus
        End If
    ElseIf Index = 8 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            chamCad_Click (0)
            LimpaControles
            txtCadastro(0).SetFocus
        End If
    ElseIf Index = 9 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCadastro(9) <> "" Then
                CarregaTipoMat
            Else
                txtCadastro(9) = ""
                txtCadastro(10) = ""
            End If
            txtCadastro(0).SetFocus
        End If
    ElseIf Index = 13 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            ChamaGridCli
        End If
    ElseIf Index = 24 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            ChamaGridCont
        End If
    ElseIf Index = 31 Then
        'If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        '    CarregaDados (index)
        'End If
    End If
End Sub

Private Sub CarregaDados(Index)
    Dim X As Integer
    If Index <> 31 Then
        If Val(txtCadastro(0)) = 0 Then
            txtCadastro(1).Enabled = True
            txtCadastro(2).Enabled = False
            txtCadastro(3).Enabled = True
            Check1.Enabled = False
            txtCadastro(5).Enabled = False
            optCadastro(1).Value = True
        
            txtCadastro(3) = "PÇ"
            txtCadastro(1) = "DIGITE O NOME DO MATERIAL"
            txtCadastro(2) = "-"
            Check1.Value = 0
            txtCadastro(1).SetFocus
            txtCadastro(1).BackColor = &HC0C0FF
            txtCadastro(3).BackColor = &HC0C0FF
            txtCadastro(4).BackColor = &HC0C0FF
            txtCadastro(8).BackColor = &HC0C0FF
            Text1.FontBold = True
            Text1.Text = "Item não cadastrado"
            
            If txtCadastro(0) = "000000" Then optCadastro_Click (1)
            Exit Sub
        Else
            txtCadastro(1).Enabled = False
            txtCadastro(2).Enabled = True
            txtCadastro(1).BackColor = &H80000005
            txtCadastro(3).BackColor = &H80000005
            txtCadastro(4).BackColor = &H80000005
            txtCadastro(8).BackColor = &H80000005
            Text1.FontBold = False
            If optCadastro(0).Value = True Then optCadastro_Click (0) Else optCadastro_Click (1)
            Check1.Enabled = True
            Check1.Value = 1
            txtCadastro(5).Enabled = True
        End If
    End If
    
    If Index = 0 Then SqlM = "Select tbMateriais.codmaterial, tbmateriais.descricao, tbMateriais.formula, tbmateriais.constpint, tbconstantes.valconst, tbmateriais.unidade, tbmateriais.forpint, tbmateriais.observacao from tbMateriais Inner Join tbconstantes on tbMateriais.codmaterial = tbConstantes.codmaterial where tbconstantes.codmaterial= '" & Val(txtCadastro(0)) & "'order by tbconstantes.codigo"
    If Index = 31 Then SqlM = "Select tbMateriais.codmaterial, tbmateriais.descricao, tbMateriais.formula, tbmateriais.constpint, tbconstantes.valconst, tbmateriais.unidade, tbmateriais.forpint, tbmateriais.observacao from tbMateriais Inner Join tbconstantes on tbMateriais.codmaterial = tbConstantes.codmaterial where tbconstantes.codmaterial= '" & Val(txtCadastro(31)) & "'order by tbconstantes.codigo"
    rsMaterial.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsMaterial.EOF Then rsMaterial.MoveFirst
    
    If Index = 0 Then rsMaterial.Find "codmaterial=" & "'" & Val(Me.txtCadastro(0)) & "'"
    If Index = 31 Then rsMaterial.Find "codmaterial=" & "'" & Val(Me.txtCadastro(31)) & "'"
    
    If rsMaterial.EOF Then
        If Index = 0 Then txtCadastro(0).Text = Format(txtCadastro(0), "000000") & ""
        If Index = 31 Then txtCadastro(0).Text = Format(txtCadastro(31), "000000") & ""
        Msgbox "Código de material não cadastrado", vbInformation, "Zeus"
    Else
        If Index = 31 Then
            txtCadastro(31).Text = Format(rsMaterial.Fields(0), "000000") & ""
            txtCadastro(32).Text = rsMaterial.Fields(1)
            Formula = rsMaterial.Fields(2)
            ForPint = rsMaterial.Fields(6)
            Text4.Text = rsMaterial.Fields(7)
        End If
        
        If Index = 0 Then
            txtCadastro(0).Text = Format(rsMaterial.Fields(0), "000000") & ""
            txtCadastro(1).Text = rsMaterial.Fields(1)
            Formula = rsMaterial.Fields(2)
            ForPint = rsMaterial.Fields(6)
            txtCadastro(3) = rsMaterial(5)
            txtCadastro(5) = rsMaterial(3)
            Text1.Text = rsMaterial.Fields(7)
            txtCadastro(3).Enabled = False
            txtCadastro(4) = 1
        End If
        For X = 1 To rsMaterial.RecordCount
            const0(X) = rsMaterial.Fields(4)
            rsMaterial.MoveNext
        Next
        If Index = 0 Then txtCadastro(4).SetFocus
    End If
    rsMaterial.Close
    Set rsMaterial = Nothing
End Sub

Private Sub CarregaTipoMat()
    Dim X As Integer
    Dim rsTipoMat As New ADODB.Recordset
    SqlM = "Select * from tbTipoMat order by tbTipoMat.codigo"
    rsTipoMat.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsTipoMat.EOF Then rsTipoMat.MoveFirst
    rsTipoMat.Find "codigo=" & "'" & Val(Me.txtCadastro(9)) & "'"
    If rsTipoMat.EOF Then
        txtCadastro(9).Text = Format(txtCadastro(9), "000000") & ""
        Msgbox "Tipo de material não cadastrado", vbInformation, "Zeus"
    Else
        txtCadastro(9).Text = Format(rsTipoMat.Fields(0), "000000") & ""
        txtCadastro(10).Text = rsTipoMat.Fields(1)
        txtCadastro(10).Enabled = False
    End If
    rsTipoMat.Close
    Set rsTipoMat = Nothing
End Sub

Private Sub CarregaFO()
    Dim rsFO As New ADODB.Recordset
    Dim rsClientes As New ADODB.Recordset
    Dim rsContatos As New ADODB.Recordset
    Dim sqlFO As String
    Dim sqlClientes As String
    Dim sqlContatos As String
    Dim rsLisview As New ADODB.Recordset
    Dim ItemLst As ListItem
    Dim sql As String
    PesoTotal = 0
    SomaTotal = 0
    SomaPint = 0
    
    sqlFO = "select * from tbfo where tbfo.codfo = '" & Val(txtCadastro(6)) & "'"
    rsFO.Open sqlFO, cnBanco, adOpenKeyset, adLockOptimistic
    If rsFO.RecordCount > 0 Then
        DTPicker1 = Format(rsFO(1), "dd/mm/yyyy")
        
        If rsFO.Fields(2) = 1 Then
            Label32 = Label32 & "Em orçamento"
        ElseIf rsFO.Fields(2) = 2 Then
            Label32 = Label32 & "Serviço"
        ElseIf rsFO.Fields(2) = 3 Then
            Label32 = Label32 & "Arquivado"
        End If
        If rsFO.Fields(9) <> "Null" Then DTPicker2 = Format(rsFO(9), "dd/mm/yyyy")
        txtCadastro(13) = Format(rsFO(5), "000000")
        txtCadastro(24) = Format(rsFO(6), "000000")
        txtCadastro(11) = rsFO(4)
        txtCadastro(28) = rsFO(7)
        txtCadastro(30) = rsFO(8)
        If rsFO.Fields(10) <> "Null" Then txtCadastro(39) = rsFO(10)
        txtCadastro(40) = Format(rsFO(11), "#,##0.000;(#,##0.000)")
        If rsFO.Fields(12) <> "Null" Then cboCadastro = rsFO(12)
        If rsFO.Fields(13) <> "Null" Then mskCadastro(1) = rsFO(13)
        If rsFO.Fields(13) <> "Null" Then mskCadastro(2) = Format(mskCadastro(1), "#,##0.000;(#,##0.000)") * Format(txtCadastro(40), "#,##0.000;(#,##0.000)")
        If rsFO.Fields(2) = 2 Then
            SkinLabel45 = Format(rsFO.Fields(3), "0000")
        Else
            SkinLabel45 = "####"
        End If
        BloqueiaControles
    Else
        LimpaControles1
        DesbloqueiaControles
        'Check2_Click
        rsFO.Close
        Set rsFO = Nothing
        Exit Sub
    End If
    
    sql = "select a.codfo,a.codseq,a.desenho,a.codmat,a.quantcj,a.dimensoes,a.pesounit,a.area,d.NOMEFANTASIA,d.CODUNDCONTROLE,a.TipoMat,a.revisao,c.descricao[DescTipoMat],a.observacao from tblistamaterial as a left join tbmateriais as b on a.codmat = b.idprd left join tbtipomat as c on a.TipoMat=c.codigo inner join " & vBancoTotvs & ".dbo.tprd as d on b.idprd = d.IDPRD where a.codfo = '" & Val(txtCadastro(6)) & "'order by a.codseq"
    rsLisview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear
    If rsLisview.RecordCount > 0 Then
        While Not rsLisview.EOF
            'insere o item do arquivo de dados
            Set ItemLst = ListView1.ListItems.Add(, , Format(rsLisview.Fields(3), "000000"))
            'cada item precisa de um subitem para exibir na lista
            ItemLst.SubItems(1) = "" & rsLisview.Fields(8)
            
            If rsLisview.Fields(10) <> 0 Then ItemLst.SubItems(2) = "" & Format(rsLisview.Fields(10), "000000") & "-" & rsLisview.Fields(12) Else ItemLst.SubItems(2) = "-"
            ItemLst.SubItems(3) = "" & rsLisview.Fields(5)
            ItemLst.SubItems(4) = "" & Format(rsLisview.Fields(6), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(5) = "" & rsLisview.Fields(9)
            ItemLst.SubItems(6) = "" & rsLisview.Fields(4)
                
            PesoTotal = rsLisview.Fields(4) * rsLisview.Fields(6)
            ItemLst.SubItems(7) = "" & Format(PesoTotal, "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(8) = "" & Format(rsLisview.Fields(7), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(9) = "" & rsLisview.Fields(2)
            'ItemLst.SubItems(10) = Format(rsLisview.Fields(3), "000000") & rsLisview.Fields(10)
            ItemLst.SubItems(11) = "" & rsLisview.Fields(11)
            
            If rsLisview.Fields(3) = 0 Then
                ItemLst.SubItems(1) = "" & rsLisview.Fields(13)
                ItemLst.SubItems(5) = "PÇ"
            End If
            If ItemLst.SubItems(2) <> "-" Then ItemLst.SubItems(10) = ItemLst.SubItems(1) & ItemLst.SubItems(2) Else ItemLst.SubItems(10) = ItemLst.SubItems(1)
            ItemLst.SubItems(12) = Format(rsLisview.Fields(1), "0000")
            SomaTotal = SomaTotal + PesoTotal
            SomaPint = SomaPint + rsLisview.Fields(7)
            PesoTotal = 0
            'vai para o proximo registro
            rsLisview.MoveNext
        Wend
    End If
    
    lblTotal = Format(SomaTotal, "#,##0.000;(#,##0.000)")
    lblTotPint = Format(SomaPint, "#,##0.000;(#,##0.000)")
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    rsLisview.Close
    ListView1.Refresh
    Set rsLisview = Nothing
    ListView2.ListItems.Clear
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView1, ColumnHeader
End Sub

Public Sub ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
    With ListView1
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
    If txtCadastro(6).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(6).Tag, vbInformation, "Atenção"
        Me.txtCadastro(6).SetFocus
        Exit Function
    End If
    If txtCadastro(30).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(30).Tag, vbInformation, "Atenção"
        Me.txtCadastro(30).SetFocus
        Exit Function
    End If
    If txtCadastro(11).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(11).Tag, vbInformation, "Atenção"
        Me.txtCadastro(11).SetFocus
        Exit Function
    End If
    If txtCadastro(13).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(13).Tag, vbInformation, "Atenção"
        Me.txtCadastro(13).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function ValidaCampo2()
    ValidaCampo2 = False
    If txtCadastro(0).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(0).Tag, vbInformation, "Atenção"
        Me.txtCadastro(0).SetFocus
        Exit Function
    End If
    If txtCadastro(4).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(4).Tag, vbInformation, "Atenção"
        Me.txtCadastro(4).SetFocus
        Exit Function
    End If
    If txtCadastro(3).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(3).Tag, vbInformation, "Atenção"
        Me.txtCadastro(3).SetFocus
        Exit Function
    End If
    
    If optCadastro(0).Value = True And txtCadastro(2).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(2).Tag, vbInformation, "Atenção"
        Me.txtCadastro(2).SetFocus
        Exit Function
    End If
    If optCadastro(1).Value = True And txtCadastro(8).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(8).Tag, vbInformation, "Atenção"
        Me.txtCadastro(8).SetFocus
        Exit Function
    End If
    If Check1.Value = 1 And txtCadastro(5).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(5).Tag, vbInformation, "Atenção"
        Me.txtCadastro(5).SetFocus
        Exit Function
    End If
    
    ValidaCampo2 = True
End Function

Private Sub CapVar()
    vAr0(Conta + 1) = Val(Mid$(txtCadastro(2), X + 2, Len(txtCadastro(2)) - X))
    Text2.Text = vAr0(Conta + 1)
    Y = X
    X = Len(txtCadastro(2))
    Conta = Conta + 1
End Sub

Private Sub CapVar2()
    vAr0(Conta + 1) = Val(Mid$(txtCadastro(34), X + 2, Len(txtCadastro(34)) - X))
    Text2.Text = vAr0(Conta + 1)
    Y = X
    X = Len(txtCadastro(34))
    Conta = Conta + 1
End Sub

Private Sub ChamaGrid()
    'Dim F As New frmpesqcli
    Sqlp = "Select * from tbmateriais where tbmateriais.descricao like '%" & txtCadastro(1) & "%'"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    F.Caption = "Pesquisa de Materiais"
    Pesquisa = frmFO.Tag
    F.Show 1
    If Pesquisa <> "0" Then
        txtCadastro(0) = Pesquisa
    End If
End Sub

Private Sub ChamaGridMat()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTipoMat order by descricao"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Tipo de Materiais"
    Pesquisa = frmFO.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "descricao=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadastro(9).Text = Format(rsLocal.Fields(0), "000000")
        Else
            Msgbox "Tipo de material não cadastrado", vbInformation, "Zeus"
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub ChamaGridCli()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbclifor order by nome"
    procnom = "nome"
    campo = 13
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa Clientes"
    Pesquisa = frmFO.Tag
    If txtCadastro(13) = "" Then F.Show 1
    Pesquisa = Mid$(Pesquisa, 7, 85)
    If Pesquisa <> "" And txtCadastro(13) = "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nome=" & "'" & Pesquisa & "'"
        If rsLocal.EOF Then
            Msgbox "Cliente não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    Else
        Sqlp = "Select * from tbclifor where tbclifor.codclifor = '" & Val(txtCadastro(13)) & "'"
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.EOF Then
            'MsgBox "Cliente não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    End If
    txtCadastro(13).Text = Format(rsLocal.Fields(0), "000000")
    txtCadastro(14).Text = rsLocal.Fields(13)
    txtCadastro(15).Text = rsLocal.Fields(1)
    txtCadastro(16).Text = rsLocal.Fields(2)
    txtCadastro(17).Text = rsLocal.Fields(3)
    txtCadastro(18).Text = rsLocal.Fields(4)
    txtCadastro(19).Text = rsLocal.Fields(5)
    txtCadastro(20).Text = Format(rsLocal.Fields(6), "(##)####-####")
    txtCadastro(21).Text = Format(rsLocal.Fields(7), "(##)####-####")
    txtCadastro(22).Text = rsLocal.Fields(8)
    txtCadastro(23).Text = rsLocal.Fields(9)
    rsLocal.Close
    Set rsLocal = Nothing
End Sub

Private Sub ChamaGridCont()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbcontatos where tbcontatos.codclifor= '" & Val(txtCadastro(13)) & "'order by nome"
    procnom = "nome"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa Contatos"
    Pesquisa = frmFO.Tag
    If txtCadastro(24) = "" Then F.Show 1
    If Pesquisa <> "" And txtCadastro(24) = "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nome=" & "'" & Mid$(Pesquisa, 7, 100) & "'"
        If rsLocal.EOF Then
            'MsgBox "Contato não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    Else
        Sqlp = "select * from tbcontatos where tbcontatos.codclifor = '" & Val(txtCadastro(13)) & "'" & _
        "and tbcontatos.codcontato=" & " '" & Val(txtCadastro(24)) & "'"
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.EOF Then
            'MsgBox "Contato não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    End If
    txtCadastro(24).Text = Format(rsLocal.Fields(1), "000000")
    txtCadastro(25).Text = rsLocal.Fields(2)
    txtCadastro(26).Text = Format(rsLocal.Fields(6), "(##)####-####")
    txtCadastro(27).Text = rsLocal.Fields(9)
    rsLocal.Close
    Set rsLocal = Nothing

End Sub

Private Sub DesbloqueiaControles()
    Dim X As Integer
    
    For X = 0 To 11
        txtCadastro(X).Enabled = True
    Next
    txtCadastro(1).Enabled = False
    txtCadastro(3).Enabled = False
    txtCadastro(13).Enabled = True
    
    mskCadastro(1).Enabled = True
    mskCadastro(2).Enabled = False
    For X = 39 To 40
        txtCadastro(X).Enabled = True
    Next
    cboCadastro.Enabled = True
    
    txtCadastro(24).Enabled = True
    txtCadastro(28).Enabled = True
    txtCadastro(29).Enabled = True
    txtCadastro(30).Enabled = True
    txtCadastro(34).Enabled = True
    DTPicker1.Enabled = True
    Check1.Enabled = True
    optCadastro(0).Enabled = True
    optCadastro(1).Enabled = True
    chamCad(0).Enabled = True
    chamCad(1).Enabled = True
    chamCad(4).Enabled = True
    chameleonButton5.Enabled = True
    chamCad(5).Enabled = True
    chameleonButton7.Enabled = True
    chameleonButton8.Enabled = True
    chameleonButton9.Enabled = True
    chamCad(3).Enabled = True
    ListView1.Enabled = True
    ListView2.Enabled = True
End Sub

Private Sub BloqueiaControles()
    txtCadastro(6).Enabled = False
End Sub

Private Sub ExportaExcel()
On Error GoTo TrataErro
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")

    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\Lista.xls"
    Plan.Visible = True
    Plan.UserControl = False

    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 10
    Me.ListView1.SortOrder = lvwAscending
    
    'PREENCHE CÉLULAS DESEJADAS
    '**********************************************************************
    Y = ListView1.ListItems.Count
    'linha1 = 27
    With Plan
        .Range("Resumo!B" & 2).Value = txtCadastro(6) ' numero FO
        .Range("Resumo!N" & 2).Value = Format(DTPicker1, "dd/mm/aaaa") ' Data FO
        .Range("Resumo!D" & 2).Value = txtCadastro(30) ' Descricao FO
        .Range("Resumo!J" & 2).Value = txtCadastro(11) ' SDC FO
        .Range("Resumo!B" & 3).Value = txtCadastro(14) ' nome cliente
        .Range("Resumo!B" & 4).Value = txtCadastro(20) ' fone cliente
        .Range("Resumo!D" & 4).Value = txtCadastro(21) 'fax cliente
        .Range("Resumo!B" & 5).Value = txtCadastro(22) 'email cliente
        .Range("Resumo!H" & 3).Value = txtCadastro(25) ' nome contato
        .Range("Resumo!H" & 4).Value = txtCadastro(26) ' fone contato
        .Range("Resumo!H" & 5).Value = txtCadastro(27) ' email contato
    End With
    
    'Range("a5", "H5").Font.Bold = True
    'Range("a5", "H5").Borders.LineStyle = 1
    'Range("a5", "H5").Borders.Weight = 3
    'Range("a5", "H5").Columns.AutoFit
    
    j = 4
    Dim valor1 As Double, valor2 As Double, valor3 As Double, valor4 As Double
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        With Plan
            valor1 = Format(ListView1.SelectedItem.ListSubItems.Item(4), "#,##0.0;(#,##0.0)")
            valor2 = Format(ListView1.SelectedItem.ListSubItems.Item(6), "#,##0.0;(#,##0.0)")
            valor3 = Format(ListView1.SelectedItem.ListSubItems.Item(7), "#,##0.0;(#,##0.0)")
            valor4 = Format(ListView1.SelectedItem.ListSubItems.Item(8), "#,##0.00;(#,##0.00)")
            
            .Range("A" & j).Value = ListView1.ListItems.Item(X)
            .Range("B" & j).Value = ListView1.SelectedItem.ListSubItems.Item(1)
            .Range("C" & j).Value = ListView1.SelectedItem.ListSubItems.Item(2)
            
            .Range("D" & j).Value = ListView1.SelectedItem.ListSubItems.Item(3)
            .Range("E" & j).Value = valor1
            .Range("F" & j).Value = ListView1.SelectedItem.ListSubItems.Item(5)
            .Range("G" & j).Value = valor2
            .Range("H" & j).Value = valor3
            .Range("I" & j).Value = valor4
            .Range("J" & j).Value = ListView1.SelectedItem.ListSubItems.Item(9)
            
            '.Range("A" & J).Borders.LineStyle = 1
            '.Range("A" & J).Borders.Weight = 2
            '.Range("B" & J).Borders.LineStyle = 1
            '.Range("B" & J).Borders.Weight = 2
            '.Range("C" & J).Borders.LineStyle = 1
            '.Range("C" & J).Borders.Weight = 2
            '.Range("D" & J).Borders.LineStyle = 1
            '.Range("D" & J).Borders.Weight = 2
            '.Range("E" & J).Borders.LineStyle = 1
            '.Range("E" & J).Borders.Weight = 2
            '.Range("F" & J).Borders.LineStyle = 1
            '.Range("F" & J).Borders.Weight = 2
            '.Range("G" & J).Borders.LineStyle = 1
            '.Range("G" & J).Borders.Weight = 2
            '.Range("H" & J).Borders.LineStyle = 1
            '.Range("H" & J).Borders.Weight = 2
            '.Range("I" & J).Borders.LineStyle = 1
            '.Range("I" & J).Borders.Weight = 2
            
            j = j + 1
        End With
    Next
    
    'Daki pra baixo eh referente ao resumo de materiais
    
    Dim compara As String
    valor3 = 0
    valor4 = 0
    j = 9
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If ListView1.SelectedItem.ListSubItems.Item(10) <> compara Then
            With Plan
                j = j + 1
                valor3 = 0
                valor4 = 0
                valor3 = Format(ListView1.SelectedItem.ListSubItems.Item(7), "#,##0.0;(#,##0.0)")
                valor4 = Format(ListView1.SelectedItem.ListSubItems.Item(8), "#,##0.00;(#,##0.00)")
                .Range("Resumo!A" & j).Value = ListView1.ListItems.Item(X)
                .Range("Resumo!B" & j).Value = ListView1.SelectedItem.ListSubItems.Item(1)
                .Range("Resumo!C" & j).Value = ListView1.SelectedItem.ListSubItems.Item(2)
                
                .Range("Resumo!E" & j).Value = ListView1.SelectedItem.ListSubItems.Item(5)
                .Range("Resumo!D" & j).Value = valor3
                .Range("Resumo!N" & j).Value = valor4
            End With
        Else
            With Plan
                valor3 = Format(valor3 + ListView1.SelectedItem.ListSubItems.Item(7), "#,##0.0;(#,##0.0)")
                valor4 = valor4 + Format(ListView1.SelectedItem.ListSubItems.Item(8), "#,##0.00;(#,##0.00)")
                .Range("Resumo!D" & j).Value = valor3
                .Range("Resumo!N" & j).Value = valor4
            End With
        End If
        compara = ListView1.SelectedItem.ListSubItems.Item(10)
    Next
    
    'Daki pra frente eh referente a lista de desenhos
    
    Me.ListView1.SortKey = 9
    Me.ListView1.SortOrder = lvwAscending
    
    j = 2
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If ListView1.SelectedItem.ListSubItems.Item(9) <> compara Then
            With Plan
                j = j + 1
                .Range("U" & j).Value = ListView1.SelectedItem.ListSubItems.Item(9)
                .Range("V" & j).Value = ListView1.SelectedItem.ListSubItems.Item(11)
            End With
        End If
        compara = ListView1.SelectedItem.ListSubItems.Item(9)
    Next
    
    Plan.Columns("A:AY").EntireColumn.AutoFit 'Ajusta as colunas

    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.Close = True
    Set Plan = Nothing
    Plan.Quit
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, O MSOffice não esta instalado nesse computador!", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub GerarResumo()
    Dim rsResumo As New ADODB.Recordset
    Dim SqlResumo As String
    
    Dim compara As String
    Dim ItemLst As ListItem
    Dim valor3 As Double, valor4 As Double, somaPL As Double, somaPM As Double
    valor3 = 0
    valor4 = 0
    somaPL = 0
    somaPM = 0
    j = 3
    Y = ListView1.ListItems.Count
    
    txtCadastro(36).Text = 5
    'optCadastro(2).Value = True
    
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 10
    Me.ListView1.SortOrder = lvwAscending
    ListView2.ListItems.Clear
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If ListView1.SelectedItem.ListSubItems.Item(10) <> compara Then
            With Plan
                j = j + 1
                valor3 = 0
                valor4 = 0
                valor3 = Format(ListView1.SelectedItem.ListSubItems.Item(7), "#,##0.00;(#,##0.00)")
                If Format(ListView1.SelectedItem.ListSubItems.Item(8), "#,##0.000;(#,##0.000)") <> "" Then valor4 = Format(ListView1.SelectedItem.ListSubItems.Item(8), "#,##0.000;(#,##0.000)")
                Set ItemLst = ListView2.ListItems.Add(, , j - 3)
                ItemLst.SubItems(1) = ListView1.ListItems.Item(X)
                ItemLst.SubItems(2) = ListView1.SelectedItem.ListSubItems.Item(1)
                ItemLst.SubItems(3) = ListView1.SelectedItem.ListSubItems.Item(2)
                ItemLst.SubItems(4) = ListView1.SelectedItem.ListSubItems.Item(5)
                ItemLst.SubItems(5) = Format(valor3, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(6) = Format(valor4, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(7) = "-"
                ItemLst.SubItems(8) = "0"
                ItemLst.SubItems(9) = "0"
                ItemLst.SubItems(10) = "1"
                ItemLst.SubItems(11) = "-"
             End With
        Else
            With Plan
                valor3 = Format(valor3 + ListView1.SelectedItem.ListSubItems.Item(7), "#,##0.000;(#,##0.000)")
                valor4 = Format(valor4 + ListView1.SelectedItem.ListSubItems.Item(8), "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(5) = Format(valor3, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(6) = Format(valor4, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(7) = "-"
                ItemLst.SubItems(8) = "0"
                ItemLst.SubItems(9) = "0"
                ItemLst.SubItems(10) = "1"
                ItemLst.SubItems(11) = "-"
            End With
        End If
        
        
         SqlResumo = "Select * from tbResumo where tbResumo.codfo= " & Val(txtCadastro(6)) & " and tbResumo.codres = '" & ItemLst & "'"
'        SqlResumo = "Select * from tbResumo where tbResumo.codfo= " & Val(txtcadastro(6)) & " and tbResumo.codmat = '" & Val(ListView1.ListItems.Item(X)) & "' and tbResumo.tipomat = '" & Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 6)) & "'"
        rsResumo.Open SqlResumo, cnBanco, adOpenKeyset, adLockOptimistic
        
        'rsResumo.Find "observacao=" & "'" & ItemLst.SubItems(2) & "'"
        If Not rsResumo.EOF Then
                'ItemLst.SubItems(4) = rsResumo.Fields(7)
                If valor3 <> rsResumo.Fields(6) Then
                    ItemLst.SubItems(11) = "Item Alterado"
                Else
                    ItemLst.SubItems(11) = "-"
                End If
                ItemLst.SubItems(7) = rsResumo.Fields(2)
                ItemLst.SubItems(8) = rsResumo.Fields(3)
                ItemLst.SubItems(9) = rsResumo.Fields(4)
                ItemLst.SubItems(10) = rsResumo.Fields(5)
        Else
            'MsgBox "aki "
            ItemLst.SubItems(11) = "Novo Item"
        End If
        rsResumo.Close
        
        somaPL = somaPL + ListView1.SelectedItem.ListSubItems.Item(7)
        somaPM = somaPM + ListView2.SelectedItem.ListSubItems.Item(8)
        
        compara = ListView1.SelectedItem.ListSubItems.Item(10)
    Next
    lbltotl = Format(somaPL, "#,##0.000;(#,##0.000)")
    lbltotpm = Format(somaPM, "#,##0.000;(#,##0.000)")
End Sub
Private Sub ResultPesq()
    txtCadastro(6) = varGlobal
End Sub

Private Sub CalcTotalProposta()
    mskCadastro(1).PromptInclude = False
    mskCadastro(2).PromptInclude = False
    If mskCadastro(1) <> "" And txtCadastro(40) <> "" Then mskCadastro(2) = Format(mskCadastro(1), "#,##0.00;(#,##0.00)") * Format(txtCadastro(40), "#,##0.00;(#,##0.00)")
    mskCadastro(1).PromptInclude = True
    mskCadastro(2).PromptInclude = True
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tbfo order by codfo Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraCodigo.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        QualForm = "novafo"
        GeraCodigo = 1
    End If
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
End Function

Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    'Para essa linha de comando existe um função dentro do módulo RotinaGeral
    'responsavel por desabilitar o BIP qdo precionada a tecla ENTER nos Texbox
    KeyAscii = Enter(KeyAscii)
    '-----------------
End Sub

Private Sub txtCadastro_LostFocus(Index As Integer)
    If Index = 40 Then
        CalcTotalProposta
    End If
End Sub

Private Sub mskCadastro_LostFocus(Index As Integer)
    If Index = 1 Then
        CalcTotalProposta
    End If
End Sub

Private Sub Mskcadastro_GotFocus(Index As Integer)
    Dim X As Integer
    For X = 1 To mskCadastro.Count - 1
        mskCadastro(X).SelStart = 0
        mskCadastro(X).SelLength = Len(mskCadastro(X).Text)
    Next
End Sub

Private Sub SomaListview()
    Dim SomaT As Currency, SomaP As Currency
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).SubItems(5) <> "pç" And ListView1.ListItems(i).SubItems(5) <> "PÇ" Then
            If ListView1.ListItems(i).SubItems(7) <> "" Then SomaT = SomaT + CCur(ListView1.ListItems(i).SubItems(7)) 'coluna de valores
            If ListView1.ListItems(i).SubItems(8) <> "" Then SomaP = SomaP + CCur(ListView1.ListItems(i).SubItems(8)) 'coluna de valores
        End If
    Next
    lblTotal.Caption = Format(SomaT, "#,##0.000;(#,##0.000)") 'Format(SomaTotal, "#,##0.000000000;(#,##0.000000000)")
    lblTotPint.Caption = Format(SomaP, "#,##0.000;(#,##0.000)")
End Sub

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadastro(6), "000000")) 'FO
        ItemLst.SubItems(1) = txtCadastro(14).Text ' Empresa
        ItemLst.SubItems(2) = txtCadastro(11).Text ' nº Coleta
        ItemLst.SubItems(3) = txtCadastro(25).Text ' Contato
        ItemLst.SubItems(4) = txtCadastro(26).Text ' Fone
        ItemLst.SubItems(5) = txtCadastro(30).Text ' Descrição
        ItemLst.SubItems(6) = DTPicker1.Value ' Data abertura
        ItemLst.SubItems(7) = DTPicker2.Value ' Data Proposta
        ItemLst.SubItems(8) = txtCadastro(39).Text ' Nº carta proposta
        ItemLst.SubItems(9) = txtCadastro(40).Text ' Quantidade
        ItemLst.SubItems(10) = Format(mskCadastro(1).Text, "R$#,##0.00;(R$#,##0.00)") ' Valor unitário
        ItemLst.SubItems(11) = Format(mskCadastro(2).Text, "R$#,##0.00;(R$#,##0.00)") ' Valor total
        ItemLst.SubItems(12) = "-"
        If SkinLabel45 = "####" Then
            ItemLst.SubItems(13) = "-" ' FCE
        Else
            ItemLst.SubItems(13) = SkinLabel45.Text ' FCE
        End If
        ItemLst.SubItems(14) = Label32
        
        If Check1.Value = 0 Then
            ItemLst.SubItems(15) = ""
            ItemLst.ListSubItems.Item(15).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(15) = ""
            ItemLst.ListSubItems.Item(15).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadastro(14).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadastro(11).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadastro(25).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtCadastro(26).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = txtCadastro(30).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = DTPicker1.Value
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = DTPicker2.Value
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = txtCadastro(39).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = txtCadastro(40).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = Format(mskCadastro(1).Text, "R$#,##0.00;(R$#,##0.00)")
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(11) = Format(mskCadastro(2).Text, "R$#,##0.00;(R$#,##0.00)")
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(12) = "-"
        If SkinLabel45 = "####" Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13) = "-"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13) = SkinLabel45.Caption
        End If
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(14) = Label32
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    Msgbox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub
