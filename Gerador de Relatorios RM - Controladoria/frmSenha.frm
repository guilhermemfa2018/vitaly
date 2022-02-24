VERSION 5.00
Begin VB.Form frmSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmSenha.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Selecione a Coligada:"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSenha.frx":1CCA
         Left            =   120
         List            =   "frmSenha.frx":1CD7
         TabIndex        =   5
         Text            =   "Vitaly"
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   720
      Picture         =   "frmSenha.frx":1CEF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "frmSenha.frx":29B9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite sua senha de acesso "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    If Combo1.Text = "Viga" Then
        vColigada = 1
    ElseIf Combo1.Text = "Vitaly" Then
        vColigada = 5
    ElseIf Combo1.Text = "Luna" Then
        vColigada = 6
    End If
End Sub

Private Sub Command1_Click()
    bot_Ok
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    vColigada = 5
    ConectaZeus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        bot_Ok ' Sub que realiza a Pesquisa no Listview mediante ao que foi digitado no TexBox Text1 e ao q foi selecionado no ComboBox Combo1
    End If
End Sub

Private Sub bot_Ok()
    
    Dim rsSenha As ADODB.Recordset
    Dim sql As String
    Set rsSenha = New ADODB.Recordset
            
    sql = "select a.codcoligada from tbsenha as a Where a.usuario= 'sergio.ob' and a.senha=" & " '" & Text1.Text & "'"
    rsSenha.Open sql, cnBancoZeus, adOpenKeyset, adLockReadOnly
    
    If Not rsSenha.EOF Then
        Form1.Command1.Visible = False
        Form1.Command3.Visible = False
        Form1.Show
        Unload Me
    ElseIf Text1 = "c41d31r4r14" Then
        Form1.Command1.Visible = True
        Form1.Command3.Visible = True
        Form1.Show
        Unload Me
    ElseIf Text1 = "yara" Then
        Form1.Command1.Visible = True
        Form1.Command3.Visible = True
        Form1.Show
        Unload Me
    ElseIf Text1 = "flaviano" Then
        Form1.Command1.Visible = True
        Form1.Command3.Visible = True
        Form1.Show
        Unload Me
    ElseIf Text1 = "vitor" Then
        Form1.Command1.Visible = True
        Form1.Command3.Visible = True
        Form1.Show
        Unload Me
    ElseIf Text1 = "flaviana" Then
        Form1.Command1.Visible = True
        Form1.Command3.Visible = True
        Form1.Show
        Unload Me
    Else
        MsgBox "Senha incorreta, tente novamente..."
    End If

    rsSenha.Close

End Sub


'ABAIXO CONEXÃO COM O BANCO DE DADOS
Private Function ConectaZeus()
'On Error GoTo Err1
    Set cnBancoZeus = New ADODB.Connection
    cnBancoZeus.Open "Provider=SQLOLEDB.1;Password=" & Form1.Text4.Text & ";Persist Security Info=True;Connect Timeout=0;User ID= " & Form1.Text5.Text & ";Initial Catalog='ZEUS';Data Source=" & Form1.Text7.Text
    Exit Function
Err1:
    Exit Function
End Function



