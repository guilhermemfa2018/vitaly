VERSION 5.00
Begin VB.Form frmAlteraSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterar Senha"
   ClientHeight    =   2055
   ClientLeft      =   5100
   ClientTop       =   4935
   ClientWidth     =   5235
   Icon            =   "frmAlteraSenha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5235
   Begin VB.CommandButton cmdCadastro 
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCadastro 
      Caption         =   "Ok"
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Confirmar Nova senha:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Nova Senha:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Senha atual:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Alterar sua senha do sistema"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmAlteraSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsSalvar As New ADODB.Recordset
Private rsLogin As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        Dim SqlSalvar As String
        Dim xcod As Integer
        SqlSalvar = "Select * from tbsenha where codcoligada ='" & vCodColigada & "' and tbsenha.usuario = '" & frmSplash.Tag & "'"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
        
        If rsSalvar.RecordCount <> 0 Then
            rsSalvar.Fields(1) = txtCadastro(1).Text
            xcod = rsSalvar.Fields(2)
                
            SqlSalvar = "Select * from tbusuarios where codcoligada ='" & vCodColigada & "' and tbusuarios.codigo = '" & xcod & "'"
            rsLogin.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
            rsLogin.Fields(12) = 0
            
            rsSalvar.Update
            rsLogin.Update
            rsSalvar.Close
            Set rsSalvar = Nothing
            rsLogin.Close
            Set rsLogin = Nothing
            MsgBox "Sua senha foi alterada com sucesso", vbInformation, "Logon"
            Unload Me
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    txtCadastro(0).Text = frmSplash.Tag
End Sub
