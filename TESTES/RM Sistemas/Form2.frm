VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificação"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdiasquefaltampararegistrar 
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtDias 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtRegistrado 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogOff 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   720
      Picture         =   "Form2.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "Form2.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogOff_Click()
    frmRegistro.Show 1
End Sub

Private Sub Command1_Click()
    If Text1.Text = senhaUsu Then
        Unload Me
        Form1.Visible = True
    Else
        MsgBox "Senha incorreta!", vbCritical, "Monitor"
    End If
End Sub

Private Sub Command2_Click()
    Unload Form2
    Form1.Visible = True
    Form1.minimize_to_tray
    Form1.WindowState = 0
    controleAcesso = 0
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    If senhaUsu = "" Then GoTo err
'**** REGISTRO *****
    If Form1.aLock.LastRunDate > Now Then
        If MsgBox("Ouve alteração na data do Sistema, inferior a data que o mesmo foi registrado " _
        & vbCrLf & "O Programa deve ser reativado na data atual ou mude a data para a data superior " _
        & vbCrLf & "que o mesmo foi registrado.", vbOKOnly + vbInformation, "Data Alterada") = vbOK Then
            DoCmd.CancelEvent
            DoCmd.Quit acQuitSaveAll
            End
        End If
    End If
    If Not Form1.aLock.RegisteredUser Then
        Me.txtdiasquefaltampararegistrar.Visible = True
    
        Dim diasQueFaltaParaRegistrar As Integer
        diasQueFaltaParaRegistrar = 0
        diasQueFaltaParaRegistrar = 30 - (Form1.aLock.UsedDays)
        Me.txtdiasquefaltampararegistrar = diasQueFaltaParaRegistrar
        cmdLogOff.Caption = "Você tem " & diasQueFaltaParaRegistrar & " dias - Registre-se"
        
        If diasQueFaltaParaRegistrar <= 0 Then
            Command1.Enabled = False
            Text1.Enabled = False
        End If
    
    Else
        Me.cmdLogOff.Visible = False
    '    Me.Caption = "SGCH ESTÁ REGISTRADO...OBRIGADO! FAVOR REINICIAR O PROGRAMA"
    '    Me.txtdiasquefaltampararegistrar.Visible = False
    '    End
    End If
    
' ********************
    
    Exit Sub
err:
    Unload Form2
    Form1.Visible = True
    Exit Sub
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        Command1_Click
    End If
End Sub
