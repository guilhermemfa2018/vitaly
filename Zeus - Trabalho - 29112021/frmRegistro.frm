VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Begin VB.Form frmRegistro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registre-se"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdregistraagora 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtcodigodoprograma 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox txtcodigoliberacao 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5040
      Width           =   3375
   End
   Begin AlphaImageControl.aicAlphaImage imgNaoLicenciado 
      Height          =   5055
      Left            =   1320
      Top             =   -600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8916
      Image           =   "frmRegistro.frx":1CCA
      Angle           =   -25
      Props           =   261
   End
   Begin AlphaImageControl.aicAlphaImage imgNaoRegistrado 
      Height          =   1200
      Left            =   480
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
      Image           =   "frmRegistro.frx":4760
      Props           =   5
   End
   Begin AlphaImageControl.aicAlphaImage imgRegistrado 
      Height          =   1200
      Left            =   3120
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
      Image           =   "frmRegistro.frx":8764
      Props           =   5
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insira aqui o código de ativação do seu sistema:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   4800
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chave do seu aplicativo:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   3975
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   5100
      Left            =   0
      Top             =   -240
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   8996
      Image           =   "frmRegistro.frx":CEB1
      Opacity         =   50
      Props           =   5
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
    
    frmSplash.aLock.LiberationKey = txtcodigoliberacao
    
    If Not frmSplash.aLock.RegisteredUser Then
        mobjMsg.Abrir "Chave de LIBERAÇÃO INCORRETA", Ok, critico, "Chave Liberação Incorreta"
        txtcodigoliberacao.SetFocus
    Else
        mobjMsg.Abrir "REGISTRO EFETUADO COM SUCESSO!", Ok, informacao, "Registro OK"
        produtoAtivado
        varGlobal = "reiniciar"
    End If

End Sub

Private Sub Form_Load()
    'If Not frmSplash.aLock.RegisteredUser Then
    '    produtoDesativado
    'Else
        produtoAtivado
    'End If
End Sub

Private Sub produtoDesativado()
    Dim diasQueFaltaParaRegistrar As Integer
    diasQueFaltaParaRegistrar = 0
    diasQueFaltaParaRegistrar = 30 - (frmSplash.aLock.UsedDays)
    lbldiasquefaltampararegistrar.ForeColor = &HC0&
    lbldiasquefaltampararegistrar = Str(diasQueFaltaParaRegistrar) & " " & lbldiasquefaltampararegistrar
    
    If diasQueFaltaParaRegistrar < 0 Then diasQueFaltaParaRegistrar = 0
    
    Label1.Visible = True
    txtcodigodoprograma.Visible = True
    Label2.Visible = True
    txtcodigoliberacao.Visible = True
    txtcodigodoprograma = frmSplash.aLock.SoftwareCode
    cmdregistraagora.Visible = True
    imgRegistrado.Visible = False
    imgNaoRegistrado.Visible = True
    imgNaoLicenciado.Visible = True
End Sub

Private Sub produtoAtivado()
   'lbldiasquefaltampararegistrar .ForeColor = &H8000&
    lbldiasquefaltampararegistrar = "Produto de propriedade da Vitaly Industria Mecânica"
    Me.Caption = "Parabéns!"
    Label1.Visible = False
    txtcodigodoprograma.Visible = False
    Label2.Visible = False
    txtcodigoliberacao.Visible = False
    cmdregistraagora.Visible = False
    imgRegistrado.Visible = True
    imgNaoRegistrado.Visible = False
    imgNaoLicenciado.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRegistro = Nothing
End Sub
