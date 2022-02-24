VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro de movimentações"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdFiltro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   1560
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
      MICON           =   "frmFiltro.frx":0CCA
      PICN            =   "frmFiltro.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdFiltro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1560
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
      MICON           =   "frmFiltro.frx":19C0
      PICN            =   "frmFiltro.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Período de Av. Eficácia "
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3255
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60293121
         CurrentDate     =   40884
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60293121
         CurrentDate     =   40884
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Configurar colunas"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro obrigatório "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmFiltro.frx":26B6
         Left            =   120
         List            =   "frmFiltro.frx":26B8
         TabIndex        =   1
         Tag             =   "Lista de opções do filtro"
         ToolTipText     =   "Lista de opções do filtro"
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFiltro_Click(Index As Integer)
On Error GoTo Err
    Select Case Index
    Case 0
        If Check1.Value = 0 Then checaFiltro = False Else checaFiltro = True
        Tipo = True
        FiltroGeral = Combo1.Text
        gravaLog "Filtro - Tipo: " & Combo1.Text, "", ""
        dataFilter1 = Format(DTPicker1.Value, "yyyy/mm/dd")
        dataFilter2 = Format(DTPicker2.Value, "yyyy/mm/dd")
        Unload Me
        Set frmFiltro = Nothing
    Case 1
        Tipo = False
        Unload Me
        Set frmFiltro = Nothing
    End Select
Err:
    Resume Next
End Sub

Private Sub Form_Load()
On Error Resume Next
    If vFil = "N" Then
        Combo1.Enabled = False
    End If
    If dataFilter1 = "" Then
        DTPicker1 = Format("01/01/" & Year(Date), "dd/mm/yyyy")
        DTPicker2 = Format("31/12/" & Year(Date), "dd/mm/yyyy")
    Else
        DTPicker1 = dataFilter1
        DTPicker2 = dataFilter2
    End If
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub
