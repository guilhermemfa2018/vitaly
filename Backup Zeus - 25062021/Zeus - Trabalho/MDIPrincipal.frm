VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000A&
   Caption         =   "SGCH - Sistema de Gestão de Competência e Habilidade"
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10560
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Picture         =   "MDIPrincipal.frx":0CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctGer 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   10560
      TabIndex        =   0
      Top             =   0
      Width           =   10560
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   6
         Left            =   4680
         TabIndex        =   10
         Tag             =   "Matriz de capacitação"
         ToolTipText     =   "Matriz de capacitação"
         Top             =   0
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
         MICON           =   "MDIPrincipal.frx":405F
         PICN            =   "MDIPrincipal.frx":407B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   5
         Left            =   3960
         TabIndex        =   9
         Tag             =   "Cursos/treinamentos"
         ToolTipText     =   "Cursos/treinamentos"
         Top             =   0
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
         MICON           =   "MDIPrincipal.frx":4D55
         PICN            =   "MDIPrincipal.frx":4D71
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   4
         Left            =   3240
         TabIndex        =   7
         Tag             =   "Cargos"
         ToolTipText     =   "Cargos"
         Top             =   0
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
         MICON           =   "MDIPrincipal.frx":5A4B
         PICN            =   "MDIPrincipal.frx":5A67
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Tag             =   "Departamentos"
         ToolTipText     =   "Departamentos"
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
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
         MICON           =   "MDIPrincipal.frx":6020
         PICN            =   "MDIPrincipal.frx":603C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   3
         Left            =   2400
         TabIndex        =   5
         Tag             =   "Setores"
         ToolTipText     =   "Setores"
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
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
         MICON           =   "MDIPrincipal.frx":65C6
         PICN            =   "MDIPrincipal.frx":65E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   7
         Left            =   5400
         TabIndex        =   1
         Tag             =   "Sair"
         ToolTipText     =   "Sair do Sistema"
         Top             =   0
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
         MICON           =   "MDIPrincipal.frx":6A06
         PICN            =   "MDIPrincipal.frx":6A22
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Tag             =   "Candidatos"
         ToolTipText     =   "Candidatos"
         Top             =   0
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
         MICON           =   "MDIPrincipal.frx":70D3
         PICN            =   "MDIPrincipal.frx":70EF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chamCad 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "Colaboradores"
         ToolTipText     =   "Colaboradores"
         Top             =   0
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
         MICON           =   "MDIPrincipal.frx":7DC9
         PICN            =   "MDIPrincipal.frx":7DE5
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7845
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Hora do sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7699
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   7710
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu mnu 
      Caption         =   "Cadastro"
      Index           =   0
      Begin VB.Menu Submnu0 
         Caption         =   "Clientes"
         Index           =   0
      End
      Begin VB.Menu Submnu0 
         Caption         =   "..."
         Index           =   1
      End
      Begin VB.Menu Submnu0 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Submnu0 
         Caption         =   "..."
         Index           =   3
      End
      Begin VB.Menu Submnu0 
         Caption         =   "..."
         Index           =   4
      End
      Begin VB.Menu Submnu0 
         Caption         =   "Tabelas Auxiliares"
         Index           =   5
         Begin VB.Menu Submnu00 
            Caption         =   "Habilidades funcionais"
            Index           =   0
         End
         Begin VB.Menu Submnu00 
            Caption         =   "Formação escolar"
            Index           =   1
         End
      End
      Begin VB.Menu Submnu0 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu Submnu0 
         Caption         =   "Usuários"
         Index           =   7
      End
      Begin VB.Menu Submnu0 
         Caption         =   "Sair"
         Index           =   8
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Movimentação"
      Index           =   1
      Begin VB.Menu Submnu1 
         Caption         =   "..."
         Index           =   0
      End
      Begin VB.Menu Submnu1 
         Caption         =   "..."
         Index           =   1
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Relatórios"
      Index           =   2
      Begin VB.Menu Submnu2 
         Caption         =   "..."
         Index           =   0
      End
      Begin VB.Menu Submnu2 
         Caption         =   "..."
         Index           =   1
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Configurações"
      Index           =   3
      Begin VB.Menu Submnu3 
         Caption         =   "Usuários"
         Index           =   0
      End
      Begin VB.Menu Submnu3 
         Caption         =   "Grupos"
         Index           =   1
      End
      Begin VB.Menu Submnu3 
         Caption         =   "Sistema"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chamCad_Click(Index As Integer)
    Select Case Index
    Case 0
        DesabBotoesN0
        Formulario = "Colaboradores"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqcolaboradores.Show 1
        HabBotoesN0
    Case 1
        DesabBotoesN0
        Formulario = "Candidatos"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqCandidatos.Show 1
        HabBotoesN0
    Case 2
        DesabBotoesN0
        Formulario = "Departamentos"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqDepartamentos.Show 1
        HabBotoesN0
    Case 3
        DesabBotoesN0
        Formulario = "Setores"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqSetores.Show 1
        HabBotoesN0
    Case 4
        DesabBotoesN0
        Formulario = "Cargos"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqCargos.Show 1
        HabBotoesN0
    Case 5
        DesabBotoesN0
        Formulario = "Treinamentos"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqTreinamentos.Show 1
        HabBotoesN0
    Case 6
        DesabBotoesN0
        Formulario = "Matriz"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqMatriz.Show 1
        HabBotoesN0
    Case 7
        DesabBotoesN0
        MDIForm_Unload 1
        HabBotoesN0
    End Select
End Sub

Private Sub MDIForm_Load()
    StatusBar1.Panels(1).Width = 1840
    StatusBar1.Panels(2).Width = 4440.189
    StatusBar1.Panels(1).Text = Format(Date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm")
    StatusBar1.Panels(2).Text = NomUsu
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Deseja encerrar a aplicação", vbQuestion + vbYesNo, "SGCG") = vbYes Then
        End
    End If
    Cancel = 1
    Exit Sub
End Sub

Private Sub Submnu0_Click(Index As Integer)
    Select Case Index
    Case 0
    Case 1
    Case 3
    Case 4
    Case 5
    Case 8
        MDIForm_Unload 1
    End Select
End Sub

Private Sub Submnu00_Click(Index As Integer)
    Select Case Index
    Case 0
        DesabBotoesN0
        Formulario = "Habilidades"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqHabilidades.Show 1
        HabBotoesN0
    Case 1
        DesabBotoesN0
        Formulario = "Escolaridade"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then frmPesqEscolar.Show 1
        HabBotoesN0
    End Select
End Sub

Private Sub Submnu1_Click(Index As Integer)
    Select Case Index
    Case 0
    Case 1
    End Select
End Sub
Private Sub chamCad_MouseOut(Index As Integer)
    Legenda = ""
    StatusBar1.Panels(3).Text = Legenda
End Sub
Private Sub chamCad_MouseOver(Index As Integer)
    Legenda = chamCad(Index).ToolTipText
    StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Submnu3_Click(Index As Integer)
    Select Case Index
    Case 0
'        frmUsuarios.Show
    Case 1
'        frmGrupos.Show
    Case 2
'        frmConfSistema.Show 1
    End Select
End Sub

