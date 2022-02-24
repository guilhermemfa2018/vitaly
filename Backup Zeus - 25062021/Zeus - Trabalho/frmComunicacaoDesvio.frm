VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComunicacaoDesvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunicação de Desvio"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComunicacaoDesvio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRNC 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gerar retrabalho?"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   7935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Revisão nº"
      Height          =   735
      Left            =   6240
      TabIndex        =   16
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtCD 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Tag             =   "nº da revisão da OS"
         ToolTipText     =   "nº da revisão da OS"
         Top             =   240
         Width           =   1815
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   1560
      OleObjectBlob   =   "frmComunicacaoDesvio.frx":0CCA
      TabIndex        =   15
      Top             =   6360
      Visible         =   0   'False
      Width           =   6735
   End
   Begin ZEUS.chameleonButton cmdCD 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   6240
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
      MICON           =   "frmComunicacaoDesvio.frx":0D24
      PICN            =   "frmComunicacaoDesvio.frx":0D40
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Comunicação "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   8175
      Begin VB.TextBox txtCD 
         Height          =   375
         Index           =   4
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   9
         Left            =   7680
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtCD 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   4
         Tag             =   "Responsável"
         Top             =   480
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmComunicacaoDesvio.frx":1A1A
         TabIndex        =   14
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtCD 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   1200
         Width           =   7935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmComunicacaoDesvio.frx":1AD2
         TabIndex        =   12
         Top             =   960
         Width           =   6855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Tag             =   "Data início"
         ToolTipText     =   "Data início"
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108920833
         CurrentDate     =   41366
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmComunicacaoDesvio.frx":1B7E
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CD nº "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtCD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Identificador"
         ToolTipText     =   "Identificador"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "OS nº "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtCD 
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Tag             =   "OS nº"
         Top             =   240
         Width           =   3735
      End
   End
   Begin ZEUS.chameleonButton cmdCD 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   6240
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
      MICON           =   "frmComunicacaoDesvio.frx":1BF8
      PICN            =   "frmComunicacaoDesvio.frx":1C14
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
Attribute VB_Name = "frmComunicacaoDesvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vEmailAprovador As String
Private rsLocal As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 9
'        ChamaGrid "tbUsuarios", "nome", txtCD(3), frmMPCompleto, "codigo", "nome"
'        txtCD(3) = Mid$(Pesquisa, 1, 6) & " - " & Mid$(Pesquisa, 7, 20)
        ChamaGridColab
        chamaChapa
    End Select
End Sub

Private Sub cmdCD_Click(Index As Integer)
    Select Case Index
    Case 0
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos e enviados com sucesso!", Ok, informacao, "ZEUS"
            If dadosEmail = False Then Exit Sub
            If Check1.Value = 1 Then 'Gerou retrabalho?
                If salvar_Dados_Retrabalho = True Then
                    mobjMsg.Abrir "RETRABALHO DISPONIVEL PARA ABERTURA DA OS!", Ok, informacao, "ZEUS"
                Else
                    mobjMsg.Abrir "ERRO NA ABERTURA DO RETRABALHO!", Ok, critico, "ZEUS"
                End If
            End If
            If vSMTP <> "" Then enviaEmail
            Unload Me
        Else
            SkinLabel1.Visible = False
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ZEUS"
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    txtCD(0).Text = Format(GeraCodigoTB("tbComunicacaoDesvio", "idcd", "", ""), "000000")
    DTPicker2 = Date
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub txtCD_GotFocus(Index As Integer)
    mudaCorText txtCD(Index)
End Sub

Private Sub txtCD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            valida_OS
        End If
    
    Case 4
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If chamaChapa = False Then Exit Sub
        End If
    End Select
End Sub

Private Sub txtCD_LostFocus(Index As Integer)
    voltaCorText txtCD(Index)
    valida_OS
End Sub

Private Function chamaChapa()
    chamaChapa = False
    Dim rschamaChapa As New ADODB.Recordset
    Dim SqlchamaChapa As String
    
    SqlchamaChapa = "select a.chapa,a.nome from " & vBancoTotvs & ".dbo.PFUNC as a where a.CODCOLIGADA = 6 and a.CODSITUACAO in('A','F','P','Z') and a.chapa = '" & Format(txtCD(4).Text, "00000") & "' UNION select a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI as chapa,a.nome COLLATE SQL_Latin1_General_CP1_CI_AI as nome from tbTerceirizados as a where a.chapa = '" & txtCD(4).Text & "' and a.ativo = 'S'"
    
    rschamaChapa.Open SqlchamaChapa, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rschamaChapa.EOF Then
        If Mid$(txtCD(4).Text, 1, 5) <> "CONTR" Then
            txtCD(4).Text = Format(txtCD(4).Text, "00000")
        End If
        txtCD(3).Text = rschamaChapa.Fields(1)  'Nome
        CompoeControles = True
    Else
        mobjMsg.Abrir "Registro de colaborador não identificado no sistema", Ok, critico, "Atenção"
        txtCD(4).Text = ""
        txtCD(3).Text = "-"
        txtCD(4).SetFocus
    End If
    rschamaChapa.Close
    Set rschamaChapa = Nothing
End Function

Private Sub ChamaGridColab()
    Dim F As New frmPesqger2
    Sqlp = "select a.chapa,a.nome from " & vBancoTotvs & ".dbo.PFUNC as a where a.CODCOLIGADA = 6 and a.CODSITUACAO in('A','F','P','Z') UNION select a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI as chapa,a.nome COLLATE SQL_Latin1_General_CP1_CI_AI as nome from tbTerceirizados as a where a.ativo = 'S'"
    procnom = "nome"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de colaboradores"
    Pesquisa = frmComunicacaoDesvio.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "chapa=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            'If Pesquisa = "Lista de Materiais" Then Pesquisa = ""
            txtCD(4) = Pesquisa
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub valida_OS()
On Error GoTo Err
    Dim rsValidaOS As New ADODB.Recordset
    Dim sqlValidaOS As String
    If IsNumeric(txtCD(1)) Then
        sqlValidaOS = "select * from tbOS where idos = '" & Val(txtCD(1)) & "' and status < 3"
        rsValidaOS.Open sqlValidaOS, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    
    If rsValidaOS.RecordCount = 0 Then
        'mobjMsg.Abrir "A OS informada não é válida ou já esta fechada", Ok, critico, "ZEUS"
        SkinLabel1.Visible = True
        SkinLabel1.Caption = "A OS informada não é válida ou já esta fechada"
    Else
        SkinLabel1.Visible = False
        txtCD(1).Text = Format(txtCD(1).Text, "000000000")
    End If
    rsValidaOS.Close
    Set rsValidaOS = Nothing
    Exit Sub
Err:
    SkinLabel1.Visible = True
    SkinLabel1.Caption = "A OS informada não é válida ou já esta fechada"
'    rsValidaOS.Close
    Set rsValidaOS = Nothing
'    mobjMsg.Abrir "A OS informada não é válida ou já esta fechada", Ok, critico, "ZEUS"
End Sub

Private Function salvar_Dados()
On Error GoTo Err
    If ValidaCampo = False Then Exit Function
    salvar_Dados = True
    txtCD(0).Text = Format(GeraCodigoTB("tbComunicacaoDesvio", "idcd", "", ""), "000000")
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(1, 1) = txtCD(0).Text
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = DTPicker2.Value 'Data de Reabertura
    vQualquerDado(2, 2) = "D"
    vQualquerDado(3, 1) = txtCD(4).Text & " - " & txtCD(3).Text
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtCD(1).Text
    vQualquerDado(4, 2) = "I"
    vQualquerDado(5, 1) = txtCD(2).Text
    vQualquerDado(5, 2) = "S"
    vQualquerDado(6, 1) = "6"
    vQualquerDado(6, 2) = "I"
    vQualquerDado(7, 1) = txtCD(5).Text
    vQualquerDado(7, 2) = "I"
    GravaDados "tbComunicacaoDesvio", "idcd", "I", txtCD(0), 7, "", "", txtCD(0)
    Exit Function
Err:
    salvar_Dados = False
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCD(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCD(1).Tag, Ok, critico, "Atenção"
        Me.txtCD(1).SetFocus
        Exit Function
    End If
    If txtCD(3).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCD(3).Tag, Ok, critico, "Atenção"
        Me.txtCD(3).SetFocus
        Exit Function
    End If
    If txtCD(5).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCD(5).Tag, Ok, critico, "Atenção"
        Me.txtCD(5).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function dadosEmail()
    dadosEmail = False
    Dim rsEnviaEmail As New ADODB.Recordset
    Dim SqlEnviaEmail As String
    SqlEnviaEmail = "Select email from tbUsuarios where codcoligada = '" & vCodcoligada & "' and nome = '" & NomUsu & "'"
    rsEnviaEmail.Open SqlEnviaEmail, cnBanco, adOpenKeyset, adLockOptimistic
    vEmailAprovador = rsEnviaEmail.Fields(0)
    If vEmailAprovador = "" Then
        mobjMsg.Abrir "Email do usuário LOGADO não está cadastrado", Ok, critico, "ZEUS"
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
    
    vDecisao = "Aprovado"
    vCorDecisao = "#228B22"

    vSMTP = "smtp.viga.ind.br"
    vUsuEmail = "taos@viga.ind.br"
    vSenhaEmail = "taos2017@"

    With Camp
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
    End With

    With Msg
        Set .Configuration = Cof
      
        .To = sEmailCD 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
'        .To = "qualidade@viga.ind.br;planejamento3@viga.ind.br;planejamento4@viga.ind.br;viga@viga.ind.br;planejamento5@viga.ind.br;planejamento7@viga.ind.br;planejamento8@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .From = vEmailAprovador  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "CD - Comunicação de Desvio nº: " & txtCD(0)
        
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
        " <B><FONT STYLE='COLOR:#009ACD'> COMUNICAÇÃO DE DESVIO </FONT></B><BR><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> A CD de nº: <b>" & txtCD(0) & "</b>, foi aberta pelo colaborador, <b>" & txtCD(3) & "</b>. Onde foi detectado que: </FONT><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & txtCD(2) & " </FONT><BR><BR><FONT STYLE='COLOR:#009ACD'> OS nº: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & txtCD(1) & "/" & txtCD(5) & " </b><BR><FONT STYLE='COLOR:#009ACD'>Data de Abertura: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & DTPicker2.Value & "</b></FONT><BR><BR>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & "" & " </FONT><BR><FONT STYLE='COLOR:#009ACD'> Att </FONT><BR><BR><BR><BR></font>" & _
        " <table border='0' cellspacing='0' width='627'>" & _
        " <tr><td width='100%'><span class='txt'>" & _
        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
        " </td></tr></table></body>"
        
        .Send
    End With
    mobjMsg.Abrir "Email enviado com suscesso!", Ok, informacao, "ZEUS"
    Exit Sub
errMail:
    Msgbox "Email não enviado para o usuário solicitante do PDO." & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "ZEUS"
    Exit Sub
End Sub

'-----GRAVA DADOS DE RETRABALHO


Private Function salvar_Dados_Retrabalho()
'On Error GoTo Err
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
        
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
        
        
    salvar_Dados_Retrabalho = True
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    txtRNC = Format(GeraCodigoTB("tbrnc", "idrnc", "", ""), "000000")
    vQualquerDado(1, 1) = txtRNC
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = txtCD(0).Text 'IDCD
    vQualquerDado(2, 2) = "I"
    vQualquerDado(3, 1) = DTPicker2.Value 'Data
    vQualquerDado(3, 2) = "D"
    vQualquerDado(4, 1) = txtCD(4) & " - " & txtCD(3).Text 'Responsável
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtCD(2).Text 'Observação
    vQualquerDado(5, 2) = "S"
    
    vQualquerDado(6, 2) = "I"
    vQualquerDado(7, 1) = "" 'Centro de Custo selecionado
    vQualquerDado(7, 2) = "S"
    vQualquerDado(8, 1) = "" 'Esboços
    vQualquerDado(8, 2) = "S"
    vQualquerDado(9, 1) = 0 'Quantidade de peças
    vQualquerDado(9, 2) = "I"
    'If DTPicker3.Value <> "" Then
    '    vQualquerDado(10, 1) = DTPicker3 'Data da re-inspeção
    '    vQualquerDado(10, 2) = "D"
    'End If
    vQualquerDado(11, 1) = "" 'Incidente
    vQualquerDado(11, 2) = "S"
    vQualquerDado(12, 1) = "" 'Correção
    vQualquerDado(12, 2) = "S"
    
    'If Check1.Value = 1 Then 'Gerou retrabalho? (1)Gerou / (0) Não gerou
        vQualquerDado(13, 1) = "S"
    'Else
    '    vQualquerDado(13, 1) = "N"
    'End If
    vQualquerDado(13, 2) = "S"
    vQualquerDado(14, 1) = "" 'Itens RNC
    vQualquerDado(14, 2) = "S"
    'If DTPicker4.Value <> "" Then
    '    vQualquerDado(15, 1) = DTPicker2.Value 'Data Conclusão
    '    vQualquerDado(15, 2) = "D"
    'End If
    
    'vQualquerDado(16, 1) = Combo1.Text 'Centro de Custo responsável pela RNC
    'vQualquerDado(16, 2) = "S"
    
    'vQualquerDado(17, 1) = txtRNC(15).Text 'Causa Raiz Determinada
    'vQualquerDado(17, 2) = "S"
    'vQualquerDado(18, 1) = txtRNC(16).Text 'Tipo de Ações - Observação
    'vQualquerDado(18, 2) = "S"
    
    'vQualquerDado(19, 1) = Combo2.Text 'Tipo de Ações - Seleção
    'vQualquerDado(19, 2) = "S"
    
    'If DTPicker5.Value <> "" Then
    '    vQualquerDado(20, 1) = DTPicker5.Value 'Data Fechamento
    '    vQualquerDado(20, 2) = "D"
    'End If
    
    GravaDados "tbRNC", "idrnc", "I", txtRNC, 20, "", "", txtRNC
        
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    
    'If Not rsSalvar.EOF Then rsSalvar.Update
    'rsSalvar.Close
    Set rsSalvar = Nothing
    Exit Function
Err:
    salvar_Dados_Retrabalho = False
End Function


