VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPDO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processo Decis�rio Organizacional"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmPDO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Decis�o "
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "Reprovar"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Tag             =   "Tomada de decis�o do PDO"
         ToolTipText     =   "Tomada de decis�o do PDO"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aprovar"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "Tomada de decis�o do PDO"
         ToolTipText     =   "Tomada de decis�o do PDO"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Observa��o "
      Height          =   1215
      Left            =   1440
      TabIndex        =   25
      Top             =   3720
      Width           =   8415
      Begin VB.TextBox txtPDO 
         Height          =   855
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Tag             =   "Observa��o"
         ToolTipText     =   "Observa��o"
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resultado"
      Height          =   975
      Left            =   7920
      TabIndex        =   23
      Top             =   2640
      Width           =   1935
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Solicita��o "
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   7695
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   975
         Index           =   6
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Avaliado "
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "CPF:"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Foto "
      Height          =   2415
      Index           =   0
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   1635
         TabIndex        =   17
         Top             =   240
         Width           =   1695
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   2175
            Left            =   0
            Top             =   -120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   3836
            Image           =   "frmPDO.frx":0CCA
         End
      End
      Begin MSComDlg.CommonDialog cdlFoto 
         Left            =   1080
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do PDO "
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtPDO 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Solicitante:"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Data da solicita��o:"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin SGCH.chameleonButton cmdPDO 
      Height          =   615
      Index           =   12
      Left            =   720
      TabIndex        =   11
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   5040
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
      MICON           =   "frmPDO.frx":0CE2
      PICN            =   "frmPDO.frx":0CFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdPDO 
      Height          =   615
      Index           =   11
      Left            =   120
      TabIndex        =   10
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   5040
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
      MICON           =   "frmPDO.frx":19D8
      PICN            =   "frmPDO.frx":19F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   1440
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   8415
   End
End
Attribute VB_Name = "frmPDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPDO As New ADODB.Recordset
Private SqlPDO As String
Private vAvaliacao As Boolean
Private vEmailAprovador As String
Private vEmailSolicitante As String

Private Sub cmdPDO_Click(Index As Integer)
    Select Case Index
    Case 11
        If MsgBox("Deseja salvar os dados do PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            gravaDadosPDO
            
            If dadosEmail = False Then Exit Sub
            If vSMTP <> "" Then enviaEmail
            Unload Me
'            gravaLog "C�digo req: " & txtCadReq(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
        End If
    Case 12
        If MsgBox("Deseja sair da tela de avalia��o do PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Unload Me
            Set frmPDO = Nothing
        End If
    End Select
End Sub

Private Sub Form_Activate()
    If vAvaliacao = False Then
        MsgBox "Essa solicita��o se encontrar processada pelo solicitante. N�o pode ser alterada"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    ResultPesq
End Sub

Private Sub ResultPesq()
    SqlPDO = "select a.id,a.datasolicitacao,a.solicitante,a.tipo,a.cpf,b.nomecolaborador,a.solicitacao,a.nota,a.decisao,a.observacao,b.foto,b.autorizacao from tbAutorizacao as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf=b.cpf where a.id = '" & Val(varGlobal) & "'"
    rsPDO.Open SqlPDO, cnBanco, adOpenKeyset, adLockReadOnly
    If IsNull(rsPDO.Fields(11)) Then vAvaliacao = False Else vAvaliacao = True
    If rsPDO.RecordCount > 0 Then
        CompoeControles
    End If
    rsPDO.Close
    Set rsPDO = Nothing
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtPDO(0).Text = Format(rsPDO.Fields(0), "000000") 'c�digo do PDO
    txtPDO(1).Text = rsPDO.Fields(1) 'data de solicita��o do PDO
    txtPDO(2).Text = rsPDO.Fields(2) 'Solicitante
    txtPDO(3).Text = rsPDO.Fields(3) 'candidato/colaborador
    txtPDO(4).Text = rsPDO.Fields(4) 'cpf do candidato/colaborador do PDO
    txtPDO(5).Text = rsPDO.Fields(5) 'nome do candidato/colaborador do PDO
    txtPDO(6).Text = rsPDO.Fields(6) 'Solicita��o do PDO
    If Not IsNull(rsPDO.Fields(9)) Then txtPDO(7).Text = rsPDO.Fields(9) 'Observa��o
    Label7 = Format(rsPDO.Fields(7), "#,##00.00;(#,##0.00)") & "%" 'Nota
    If IsNull(rsPDO.Fields(8)) Then
        Option1.Value = False
        Option2.Value = False
    ElseIf Trim(rsPDO.Fields(8)) = "Aprovado" Then
        Option1.Value = True
        Option2.Value = False
    ElseIf Trim(rsPDO.Fields(8)) = "Reprovado" Then
        Option1.Value = False
        Option2.Value = True
    End If
    Label53.Caption = rsPDO.Fields(10)
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    
    If RemoveMask(Val(Label7)) < MediaGlobal And RemoveMask(Val(Label7)) >= vAprovadoRest Then
        Label7.ForeColor = &H40C0&
    ElseIf Val(Label7) < vAprovadoRest Then
        Label7.ForeColor = &HC0&
    ElseIf Val(Label7) >= MediaGlobal Then
        Label7.ForeColor = &H8000&
    End If
End Sub

Private Sub gravaDadosPDO()
    If ValidaCampos = False Then Exit Sub
    Dim rsSalvarPDO As New ADODB.Recordset
    Dim SqlSalvarPDO As String
    Dim vDecisao As String
    If Option1.Value = True Then
        vDecisao = "Aprovado"
    Else
        vDecisao = "Reprovado"
    End If
    SqlSalvarPDO = "Update tbAutorizacao set aprovador = '" & NomUsu & "', datadecisao = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))), observacao = '" & txtPDO(7) & "', status = 'S', decisao = '" & Trim(vDecisao) & " ' Where codcoligada = '" & vCodcoligada & "' and id = '" & Val(txtPDO(0)) & "'"
    rsSalvarPDO.Open SqlSalvarPDO, cnBanco
    MsgBox "Os dados do PDO foram salvos com sucesso", vbInformation, "SGCH"
    AtualizaListview
    Exit Sub
End Sub

Private Function dadosEmail()
    dadosEmail = False
    Dim rsEnviaEmail As New ADODB.Recordset
    Dim SqlEnviaEmail As String
    SqlEnviaEmail = "Select email from tbUsuarios where codcoligada = '" & vCodcoligada & "' and nome = '" & NomUsu & "'"
    rsEnviaEmail.Open SqlEnviaEmail, cnBanco, adOpenKeyset, adLockOptimistic
    vEmailAprovador = rsEnviaEmail.Fields(0)
    If vEmailAprovador = "" Then
        MsgBox "Email do usu�rio LOGADO n�o est� cadastrado"
        Exit Function
    End If
    rsEnviaEmail.Close
    SqlEnviaEmail = "Select email from tbUsuarios where codcoligada = '" & vCodcoligada & "' and nome = '" & txtPDO(2) & "'"
    rsEnviaEmail.Open SqlEnviaEmail, cnBanco, adOpenKeyset, adLockOptimistic
    vEmailSolicitante = rsEnviaEmail.Fields(0)
    If vEmailSolicitante = "" Then
        MsgBox "Email do usu�rio SOLICITANTE n�o est� cadastrado"
        Exit Function
    End If
    rsEnviaEmail.Close
    Set rsEnviaEmail = Nothing
    dadosEmail = True
End Function

Private Sub enviaEmail()
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    If Option1.Value = True Then
        vDecisao = "Aprovado"
        vCorDecisao = "#228B22"
    Else
        vDecisao = "Reprovado"
        vCorDecisao = "#CD2626"
    End If

    With Camp
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   �informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autentica��o
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
    End With

    With Msg
        Set .Configuration = Cof
      
        .To = vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br � destinatarios separados por ;
        .From = vEmailAprovador  '"contatos@flowsys.com.br"   'remetente@email.com.br � remetente"
        .Subject = "SGCH - Resposta PDO n�: " & txtPDO(0)
        
        .HTMLBody = "<html>" & _
        " <head>" & _
        " <meta http-equiv='Content-Type'" & _
        " content='text/html; charset=iso-8859-1'>" & _
        " <meta name='GENERATOR' content='Microsoft FrontPage Express 2.0'>" & _
        " <title>Assinatura</title>" & _
        " <STYLE type='text/css'>" & _
        " <!-- -->" & _
        " </STYLE></head>" & _
        " <body link='#0000FF' vlink='#800080'>" & _
        " <font face = 'Courier New' size = 5>" & _
        " <B><FONT STYLE='COLOR:#009ACD'> SERVI�O DE EMAIL SGCH </FONT></B><BR><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> O PDO de n�: <b>" & txtPDO(0) & "</b>, referente ao " & txtPDO(3) & ", <b>" & txtPDO(5) & "</b>. Onde foi detectado que: </FONT><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & txtPDO(6) & " </FONT><BR><BR><FONT STYLE='COLOR:#009ACD'> Pontua��o: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & Label7 & " </b><BR><FONT STYLE='COLOR:#009ACD'>Status decis�rio: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & UCase(vDecisao) & "</b></FONT><BR><BR>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & txtPDO(7) & " </FONT><BR><FONT STYLE='COLOR:#009ACD'> Att </FONT><BR><BR><BR><BR></font>" & _
        " <table border='0' cellspacing='0' width='627'>" & _
        " <tr><td width='100%'><span class='txt'>" & _
        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
        " </td></tr></table></body>"
        
'        .HTMLBody = "<html>" & _
'        " <head>" & _
'        " <meta http-equiv='Content-Type'" & _
'        " content='text/html; charset=iso-8859-1'>" & _
'        " <meta name='GENERATOR' content='Microsoft FrontPage Express 2.0'>" & _
'        " <title>Assinatura</title>" & _
'        " <STYLE type='text/css'>" & _
'        " <!-- -->" & _
'        " </STYLE></head>" & _
'        " <body link='#0000FF' vlink='#800080'>" & _
'        " <font face = 'Courier New' size = 5>" & _
'        " <B><FONT STYLE='COLOR:#009ACD'> SERVI�O DE EMAIL SGCH </FONT></B><BR><BR></font>" & _
'        " <font face = 'Courier New' size = 2>" & _
'        " <FONT STYLE='COLOR:#009ACD'> O PDO de n�: <b>" & txtPDO(0) & "</b>, referente � admiss�o do " & txtPDO(3) & ", <b>" & txtPDO(5) & "</b> foi <b>" & UCase(vDecisao) & "</b> </FONT><BR>" & _
'        " </font>" & _
'        " <font face = 'Courier New' size = 2>" & _
'        " <FONT STYLE='COLOR:#009ACD'> " & txtPDO(7) & " </FONT><BR><FONT STYLE='COLOR:#009ACD'> Att </FONT><BR><BR><BR><BR></font>" & _
'        " <table border='0' cellspacing='0' width='627'>" & _
'        " <tr><td width='100%'><span class='txt'>" & _
'        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
'        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
'        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
'        " </td></tr></table></body>"
        
        
        '.HTMLBody = "<b> teste </b> de evio de email" 'strHTML
        '.CC = 'Informe o ou os destinat�rios da c�pia
        '.BCC = "contatos@flowsys.com.br"   'Informe o ou os destinat�rios da c�pia oculta
        '.AddAttachment �c:    este1.txt;c:    este2.txt� ' informe o ou os anexos


        .Send
    End With
    MsgBox "Email enviado com suscesso!"
    Exit Sub
errMail:
    MsgBox "Email n�o enviado para o usu�rio solicitante do PDO." & vbCrLf & vbCrLf & _
    "ERRO de autentica��o! Favor verificar se as configura��es de SMTP e email est�o corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SGCH"
    Exit Sub
End Sub

Private Function ValidaCampos()
    ValidaCampo = False
    If Option1.Value = False And Option2.Value = False Then
        MsgBox "Favor informar o campo " & Option1.Tag, vbCritical, "Aten��o"
        Me.txtPDO(7).SetFocus
        Exit Function
    End If
    ValidaCampos = True
End Function

Private Sub AtualizaListview()
'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Option1.Value = True Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = "Aprovado"
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = "Reprovado"
    End If
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ReportIcon = "OK"
    Exit Sub
Err:
    MsgBox "N�o foi poss�vel realizar as altera��es", vbInformation, "Aten��o"
    Exit Sub
End Sub

