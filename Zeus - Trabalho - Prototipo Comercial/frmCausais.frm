VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmCausais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Causais"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCausais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12938
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
      TabCaption(0)   =   "Grupos"
      TabPicture(0)   =   "frmCausais.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Causais"
      TabPicture(1)   =   "frmCausais.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Grupo"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Left            =   8880
            TabIndex        =   28
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtCausal 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtCausal 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   8
            Top             =   480
            Width           =   7575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmCausais.frx":0D02
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmCausais.frx":0D6E
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Causal"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   9375
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   7
            Left            =   1920
            TabIndex        =   14
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
            MICON           =   "frmCausais.frx":0DD4
            PICN            =   "frmCausais.frx":0DF0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   6
            Left            =   1320
            TabIndex        =   13
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
            MICON           =   "frmCausais.frx":1ACA
            PICN            =   "frmCausais.frx":1AE6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   5
            Left            =   720
            TabIndex        =   12
            Top             =   960
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
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
            MICON           =   "frmCausais.frx":27C0
            PICN            =   "frmCausais.frx":27DC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtCausal 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   975
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   3975
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7011
            LabelEdit       =   1
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
         Begin VB.TextBox txtCausal 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   10
            Top             =   480
            Width           =   8055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmCausais.frx":34B6
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmCausais.frx":3522
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   4
            Left            =   120
            TabIndex        =   11
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
            MICON           =   "frmCausais.frx":3588
            PICN            =   "frmCausais.frx":35A4
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
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Processo "
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
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   9375
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   3
            Left            =   1920
            TabIndex        =   5
            Top             =   840
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
            MICON           =   "frmCausais.frx":427E
            PICN            =   "frmCausais.frx":429A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   2
            Left            =   1320
            TabIndex        =   4
            Top             =   840
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
            MICON           =   "frmCausais.frx":4F74
            PICN            =   "frmCausais.frx":4F90
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   1
            Left            =   720
            TabIndex        =   3
            Top             =   840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
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
            MICON           =   "frmCausais.frx":5C6A
            PICN            =   "frmCausais.frx":5C86
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtCausal 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCausal 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   1
            Top             =   480
            Width           =   8175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "frmCausais.frx":6960
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmCausais.frx":69CC
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   840
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
            MICON           =   "frmCausais.frx":6A32
            PICN            =   "frmCausais.frx":6A4E
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
            Height          =   5055
            Left            =   120
            TabIndex        =   6
            Top             =   1560
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   8916
            LabelEdit       =   1
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
      End
   End
   Begin ZEUS.chameleonButton cmdcadastro 
      Height          =   615
      Index           =   11
      Left            =   720
      TabIndex        =   17
      Top             =   7560
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
      MICON           =   "frmCausais.frx":7728
      PICN            =   "frmCausais.frx":7744
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdcadastro 
      Height          =   615
      Index           =   12
      Left            =   120
      TabIndex        =   16
      Top             =   7560
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
      MICON           =   "frmCausais.frx":841E
      PICN            =   "frmCausais.frx":843A
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
Attribute VB_Name = "frmCausais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsLocal As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        'Inclusão - Listview1
        If ValidaCampos(ListView1, txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)) = False Then Exit Sub
        IncluirLV ListView1, txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
        LimpaControles txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
        txtCausal(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 1
        'Novo - Listview1
        LimpaControles txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
        txtCausal(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 2
        'Edição - Listview1
        AlteraLV ListView1, txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
    Case 3
        'Exclusão - Listview1
        ExcluirItemLV ListView1
        LimpaControles txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
        txtCausal(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 4
        'Inclusão - Listview2
        If ValidaCampos(ListView2, txtCausal(4), txtCausal(5), txtCausal(2), txtCausal(3), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4)) = False Then Exit Sub
        IncluirLV ListView2, txtCausal(4), txtCausal(5), txtCausal(2), txtCausal(3), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4)
        LimpaControles txtCausal(4), txtCausal(5), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4)
        txtCausal(4) = Format(GeraCodigoLV(ListView2), "00")
    Case 5
        'Novo - Listview2
        LimpaControles txtCausal(4), txtCausal(5), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4)
        txtCausal(4) = Format(GeraCodigoLV(ListView2), "00")
    Case 6
        'Edição - Listview2
        AlteraLV ListView2, txtCausal(4), txtCausal(5), txtCausal(2), txtCausal(3), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
    Case 7
        'Exclusão - Listview1
        ExcluirItemLV ListView2
        LimpaControles txtCausal(4), txtCausal(5), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4)
        txtCausal(4) = Format(GeraCodigoLV(ListView2), "00")
    Case 12
        'Gravação -  ListView1
        limpaQualquerDado
        ordenaLVArray ListView1, "0", "1", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbcausaisgrupos", "", "I", txtCausal(0)
        
        'Gravação -  ListView2
        limpaQualquerDado
        ordenaLVArray ListView2, "0", "1", "2", "", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbcausais", "", "I", txtCausal(0)
        
        mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
    Case 11
        Unload Me
    End Select
End Sub

Private Sub Command1_Click()
    ChamaGridGrupoCausais
    CarregaTxt "tbCausaisGrupos", "idgrupocausal", "I", "", "", txtCausal(2), txtCausal(2), 0, 1, txtCausal(2), "I", txtCausal(3), "1"
    LimpaControles txtCausal(4), txtCausal(5), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4), txtCausal(4)
    txtCausal(4) = Format(GeraCodigoLV(ListView2), "00")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    SSTab1.Tab = 0
    listview_cabecalho
    chamaSQL "Select a.idgrupocausal,a.nomegrupocausal from tbCausaisGrupos as a Order by a.idgrupocausal"
    Compoe_Listview ListView1, Sqlp, "00"
    
    chamaSQL "Select a.idcausal,a.nomecausal,a.idgrupocausal,b.nomegrupocausal from tbCausais as a inner join tbcausaisgrupos as b on a.idgrupocausal = b.idgrupocausal Order by a.idgrupocausal"
    Compoe_Listview ListView2, Sqlp, "00"
    
    LimpaControles txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
    txtCausal(0) = Format(GeraCodigoLV(ListView1), "00")
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
    ListView1.ColumnHeaders.Add , , "ID Grupo", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Grupo", ListView1.Width / 1.5
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID Causal", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Causal", ListView2.Width / 1.5
    ListView2.ColumnHeaders.Add , , "ID Grupo", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Grupo", ListView2.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ChamaGridGrupoCausais()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "Select a.idgrupocausal,a.nomegrupocausal from tbcausaisgrupos as a order by a.nomegrupocausal"
    procnom = "nomegrupocausal"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Grupos de Causais"
    Pesquisa = frmCausais.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "idgrupocausal=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            'If Pesquisa = "Lista de Materiais" Then Pesquisa = ""
            txtCausal(2) = Pesquisa
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub ListView1_DblClick()
    AlteraLV ListView1, txtCausal(0), txtCausal(1), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
End Sub

Private Sub ListView2_DblClick()
    AlteraLV ListView2, txtCausal(4), txtCausal(5), txtCausal(2), txtCausal(3), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0), txtCausal(0)
End Sub

Private Sub txtCausal_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtCausal(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtCausal.Count - 1
        txtCausal(X).SelStart = 0
        txtCausal(X).SelLength = Len(txtCausal(X).Text)
    Next
End Sub

Private Sub txtCausal_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaTxt "tbCausaisGrupos", "idgrupocausal", "I", "", "", txtCausal(2), txtCausal(2), 0, 1, txtCausal(2), "I", txtCausal(3), "1"
            txtCausal(4) = Format(GeraCodigoLV(ListView2), "00")
        End If
    End Select
End Sub

Private Sub txtCausal_LostFocus(Index As Integer)
    voltaCorText txtCausal(Index)
    Select Case Index
    Case 2
        Pesquisa = 1
        CarregaTxt "tbCausaisGrupos", "idgrupocausal", "I", "", "", txtCausal(2), txtCausal(2), 0, 1, txtCausal(2), "I", txtCausal(3), "1"
        txtCausal(4) = Format(GeraCodigoLV(ListView2), "00")
    End Select
End Sub
