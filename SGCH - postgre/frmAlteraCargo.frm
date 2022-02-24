VERSION 5.00
Begin VB.Form frmAlteraCargo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração funcional abaixo do tempo mínimo permitido"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   Icon            =   "frmAlteraCargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Justifique a alteração do cargo "
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmAlteraCargo.frx":0CCA
         Left            =   120
         List            =   "frmAlteraCargo.frx":0CD1
         TabIndex        =   3
         Top             =   360
         Width           =   6135
      End
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
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
      MICON           =   "frmAlteraCargo.frx":0CF0
      PICN            =   "frmAlteraCargo.frx":0D0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "Confirmar"
      ToolTipText     =   "Confirmar"
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
      MICON           =   "frmAlteraCargo.frx":19E6
      PICN            =   "frmAlteraCargo.frx":1A02
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
Attribute VB_Name = "frmAlteraCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNovoCol_Click(Index As Integer)
    If Combo1.Text = "" Then
        MsgBox "Favor justificar a alteração funcional", vbCritical, "SGCH"
        Exit Sub
    End If
    Select Case Index
    Case 0
        If MsgBox("Confirma a justificativa da alteração funcional?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            vsituacao = Combo1.Text
            'gravaLog "CPF: " & txtNovoColaborador(0) & ", Registro: " & txtNovoCol(1), "Nome: " & txtNovoColaborador(1), "Média Geral: " & Label41 & ", Status: " & Label9
            Unload Me
        End If
    Case 2
        vsituacao = ""
        Unload Me
        Set frmAlteraCargo = Nothing
    End Select
End Sub

Private Sub Form_Load()
    vsituacao = ""
End Sub
