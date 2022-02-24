VERSION 5.00
Begin VB.Form frmPrinterProg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios da programação"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "frmPrinterProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdImprimir 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   2880
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
      MPTR            =   0
      MICON           =   "frmPrinterProg.frx":3469A
      PICN            =   "frmPrinterProg.frx":346B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdImprimir 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2880
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
      MPTR            =   0
      MICON           =   "frmPrinterProg.frx":35390
      PICN            =   "frmPrinterProg.frx":353AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optImprimir 
      Caption         =   "Avaliação do treinamento"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3855
   End
   Begin VB.OptionButton optImprimir 
      Caption         =   "Certificado de treinamento"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   3735
   End
   Begin VB.OptionButton optImprimir 
      Caption         =   "Avaliação de eficácia do treinamento (individual)"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
   Begin VB.OptionButton optImprimir 
      Caption         =   "Avaliação de eficácia do treinamento (geral)"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.OptionButton optImprimir 
      Caption         =   "Lista de presença do treinamento"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmPrinterProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimir_Click(Index As Integer)
    Select Case Index
    Case 0
        If optImprimir(0) Then FCRListaPresenca.Show 1
        If optImprimir(1) Then FCRAET.Show 1
        If optImprimir(2) Then FCRAvalDesempenho.Show 1
        If optImprimir(3) Then frmCertificado.Show 1
        If optImprimir(4) Then FCRAvTrei.Show 1
    Case 1
        Unload Me
        Set frmPrinterProg = Nothing
    End Select
End Sub

Private Sub Form_Load()
    If chamaForm.txtProgTrei(0) = "-" Then
        optImprimir(0).Enabled = False
        optImprimir(1).Enabled = False
        optImprimir(2).Enabled = False
        optImprimir(3).Enabled = False
    End If
    
    If chamaForm.Check1.Value = 0 Then
        optImprimir(1).Enabled = False
        optImprimir(2).Enabled = False
    Else
        If StatusTrei = "Concluido" Then
            'optImprimir(1).Enabled = True
            'optImprimir(3).Enabled = True
            optImprimir(0).Enabled = True
            optImprimir(1).Enabled = True
            optImprimir(2).Enabled = True
            optImprimir(3).Enabled = True
        Else
            optImprimir(1).Enabled = False
            'optImprimir(2).Enabled = False
            optImprimir(3).Enabled = False
        End If
    End If
    configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub configControles()
    If vImp = "N" Then
        cmdImprimir(0).UseGreyscale = True
        cmdImprimir(0).DragMode = 1
        cmdImprimir(0).SpecialEffect = cbEngraved
    End If
End Sub
