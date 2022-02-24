VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Begin VB.Form frmConfirmaChapa 
   BackColor       =   &H80000009&
   Caption         =   "Confirmação"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfirmaChapa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   " Favor passar o cracha de identificação novamente"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4800
         Top             =   4920
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "-"
         Top             =   4440
         Width           =   975
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
         Height          =   5460
         Left            =   960
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   9631
         Image           =   "frmConfirmaChapa.frx":0CCA
         Props           =   5
      End
   End
End
Attribute VB_Name = "frmConfirmaChapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private temp As String ' variavel global que pega os valores de entrada do leitor
Private vCBarraGeral As String
Private vTempo As Integer

Private Function validaCracha()
On Error GoTo Err
    Dim rsAchaCC As New ADODB.Recordset
    Dim SqlAchaCC As String
    
    Dim rsValidaCracha As New ADODB.Recordset
    Dim SqlvalidaCracha As String
    
    validaCracha = False
    
    vCBarraGeral = ""
    
    'Text1.Text = Format(Val(Text1.Text), "00000")
    If vVerificaPermissao = 0 Then
        SqlAchaCC = "Select a.codigobarra,b.idos,b.idoperacao,b.idcc from tbOsMov as a inner join tbmpitens as b on a.codigobarra = b.codigobarra " & _
                    "where a.chapa = '" & Text1.Text & "' and datasai is null"
    ElseIf vVerificaPermissao = 1 Then
        SqlAchaCC = "Select a.codigobarra,a.idos,a.idoperacao,a.idcc from tbmpitens as a where a.codigobarra = '" & Form1.txtOS(1).Text & "'"
    End If
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    
    If rsAchaCC.RecordCount = 0 Then
        validaCracha = False
        Dim rsAchaNome As New ADODB.Recordset
        Dim SqlAchaNome As String
        If Mid$(Text1.Text, 1, 5) = "CONTR" Or Mid$(Text1.Text, 1, 5) = "contr" Then
            SqlAchaNome = "select a.nome from tbTerceirizados as a where a.chapa = '" & Text1.Text & "'"
        Else
            SqlAchaNome = "select b.NOME from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODPESSOA = b.CODIGO where a.CHAPA = '" & Format(Text1.Text, "00000") & "'"
        End If
        rsAchaNome.Open SqlAchaNome, cnBanco, adOpenKeyset, adLockReadOnly
        vNomeGlobal = rsAchaNome.Fields(0)
        rsAchaNome.Close
        Set rsAchaNome = Nothing
        
        rsAchaCC.Close
        Set rsAchaCC = Nothing
        Exit Function
    Else
        'Dim rsAchaNome As New ADODB.Recordset
        'Dim SqlAchaNome As String
        If Mid$(Text1.Text, 1, 5) = "CONTR" Then
            SqlAchaNome = "select a.nome from tbTerceirizados as a where a.chapa = '" & Text1.Text & "'"
        Else
            SqlAchaNome = "select b.NOME from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODPESSOA = b.CODIGO where a.CHAPA = '" & Format(Text1.Text, "00000") & "'"
        End If
        rsAchaNome.Open SqlAchaNome, cnBanco, adOpenKeyset, adLockReadOnly
        vNomeGlobal = rsAchaNome.Fields(0)
        rsAchaNome.Close
        Set rsAchaNome = Nothing
    End If
    
    vCBarraGeral = rsAchaCC.Fields(0)
    
    SqlvalidaCracha = "select * from tbautCCusto where chapa = '" & Text1.Text & "' and idcc = '" & rsAchaCC.Fields(3) & "'"
    rsValidaCracha.Open SqlvalidaCracha, cnBanco, adOpenKeyset, adLockReadOnly
    If rsValidaCracha.RecordCount > 0 Then
        validaCracha = True
    Else
        vTempo = 2
    End If
    
        If Form1.txtOS(1).Text = "" Then
            validaCracha = True
            vTempo = 0
        End If
    
    
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    
    rsValidaCracha.Close
    Set rsValidaCracha = Nothing
    
    Exit Function
Err:
    validaCracha = False
    Exit Function
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Err
    Timer1.Enabled = True
    If KeyAscii <> 13 Then
        temp = temp & Chr(KeyAscii)
    Else
        Text1.Text = ""
        Text1.Text = temp
        If validaCracha = False And vTempo > 1 Or validaCracha = True And vTempo > 1 Then
            vMSGGlobal = "NO"
            zeraTempo
        Else
            vMSGGlobal = "OK"
        End If
        Unload Me
    End If
    Exit Sub
Err:
    MsgBox "Por favor passe o CRACHA no leitor de Código de Barras"
End Sub

Private Sub Form_Load()
    AlwaysOnTop Me, True ' Mantem o formulário sempre em primeiro plano
    temp = ""
    Text1.Text = ""
    vNomeGlobal = ""
    vTempo = 0
End Sub

Private Sub zeraTempo()
    Timer1.Enabled = False
    vTempo = 0
    Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
    vTempo = vTempo + 1
End Sub
