VERSION 5.00
Begin VB.Form frmRegistro 
   Caption         =   "Formulário de registro"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   Icon            =   "frmRegistro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcodigoliberacao 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox txtcodigodoprograma 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtdiasquefaltampararegistrar 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   -120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Comando45 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdregistraagora 
      Caption         =   "Registrar agora"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Insira aqui o código de ativação do seu sistema:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Chave do seu aplicativo:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Dias restantes de uso:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   -120
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdregistraagora_Click()
    If txtcodigoliberacao = "" Then
        txtcodigoliberacao.SetFocus
        Exit Sub
    End If
    
    Form1.aLock.LiberationKey = txtcodigoliberacao
    
    If Not Form1.aLock.RegisteredUser Then
        MsgBox "Chave de LIBERAÇÃO INCORRETA", vbOKOnly + vbCritical, "Chave Liberação Incorreta"
        txtcodigoliberacao.SetFocus
    Else
        MsgBox "REGISTRO EFETUADO COM SUCESSO !", vbExclamation, "Registro OK"
        Form2.Caption = "VERSÃO REGISTRADA"
        Form2.txtDias.Visible = False
        Form2.txtdiasquefaltampararegistrar.Visible = False
        Form2.txtRegistrado.Visible = False
        
        MsgBox "A aplicação será fechada. Inicie-a novamente"
        End
    End If
End Sub

Private Sub cmdregistrardepois_Click()

End Sub

Private Sub Comando45_Click()
    End
End Sub

Private Sub Form_Load()
    Dim diasQueFaltaParaRegistrar As Integer
    diasQueFaltaParaRegistrar = 0
    diasQueFaltaParaRegistrar = 30 - (Form1.aLock.UsedDays)
    txtdiasquefaltampararegistrar = diasQueFaltaParaRegistrar

    'If diasQueFaltaParaRegistrar <= 0 Then
        'cmdregistrardepois.Enabled = False
    'End If
    
    txtcodigodoprograma = Form1.aLock.SoftwareCode
End Sub
