VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Composição dos critérios dos Grupos"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   17295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame21 
      Caption         =   "Composição da notas do indicadores "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   5520
      TabIndex        =   12
      Top             =   120
      Width           =   11655
      Begin SAF.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   7
         Left            =   1920
         TabIndex        =   22
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
         MICON           =   "Form1.frx":0CCA
         PICN            =   "Form1.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SAF.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   8
         Left            =   1320
         TabIndex        =   23
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
         MICON           =   "Form1.frx":19C0
         PICN            =   "Form1.frx":19DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SAF.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   9
         Left            =   720
         TabIndex        =   24
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
         MICON           =   "Form1.frx":26B6
         PICN            =   "Form1.frx":26D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtNota 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   9015
      End
      Begin VB.TextBox txtNota 
         Height          =   330
         Index           =   1
         Left            =   9240
         TabIndex        =   17
         Top             =   480
         Width           =   2295
      End
      Begin VB.Frame Frame22 
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
         Height          =   735
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   1695
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "-"
            Top             =   240
            Width           =   1455
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "Form1.frx":33AC
         TabIndex        =   13
         Top             =   1080
         Width           =   6015
      End
      Begin SAF.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   11
         Left            =   120
         TabIndex        =   14
         Tag             =   "Salvar Notas"
         ToolTipText     =   "Salvar Notas"
         Top             =   7440
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
         MICON           =   "Form1.frx":3404
         PICN            =   "Form1.frx":3420
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   9240
         OleObjectBlob   =   "Form1.frx":40FA
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Form1.frx":417E
         TabIndex        =   20
         Top             =   240
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9975
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SAF.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   25
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
         MICON           =   "Form1.frx":41E8
         PICN            =   "Form1.frx":4204
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
   Begin VB.Frame Frame18 
      Caption         =   "Grupos "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtValida 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtValida 
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame19 
      Caption         =   "Critérios "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5295
      Begin VB.TextBox txtValida 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtValida 
         Height          =   375
         Index           =   4
         Left            =   2880
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3836
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame20 
      Caption         =   "Sub-critérios "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   5295
      Begin VB.TextBox txtValida 
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtValida 
         Height          =   375
         Index           =   5
         Left            =   2880
         TabIndex        =   1
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4683
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin SAF.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   8400
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
      MICON           =   "Form1.frx":4EDE
      PICN            =   "Form1.frx":4EFA
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    listview_cabecalho_Nota
    chamaSQL "select a.idgrupocriterio,a.nomegrupocriterio from tbGrupoCriterio as a"
    Compoe_Listview ListView1, Sqlp, "00"

    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico

End Sub
