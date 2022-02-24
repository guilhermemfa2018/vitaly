VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmInformaQtd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe a quantidade"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "frmInformaQtd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   1560
      Picture         =   "frmInformaQtd.frx":0CCA
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtInforma 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "frmInformaQtd.frx":1994
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtInforma 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmInformaQtd.frx":1A0C
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmInformaQtd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    VerificaQtd
End Sub

Private Sub Form_Activate()
    txtInforma(0).SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    txtInforma(1).Text = vQtdDisponivel
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    Me.Top = (Principal.Height / 2) - (Me.Height / 2)
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub txtInforma_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            VerificaQtd
        End If
    End Select
End Sub

Private Sub txtInforma_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 0
        'aceitar somente números e "Back Space", "Enter", "virgula"
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 And KeyAscii <> 45 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub VerificaQtd()
On Error GoTo Err
    If Val(txtInforma(0).Text) > Val(txtInforma(1).Text) Then
        Msgbox "Quantidade informada maior que a quantidade disponível"
        vQtdSolicitada = 0
    Else
        vQtdSolicitada = txtInforma(0).Text
        Unload Me
    End If
    Exit Sub
Err:
    Msgbox "Nenhuma quantidade foi especificada"
End Sub
