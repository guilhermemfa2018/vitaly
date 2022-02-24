VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmAlteraStatusFCE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aterar Status da FCE"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmAlteraStatusFCE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmAlteraStatusFCE.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmAlteraStatusFCE.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "Salvar Status"
      ToolTipText     =   "Salvar Status"
      Top             =   4080
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione o novo status da FCE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   7935
      Begin VB.Frame Frame3 
         Caption         =   "Confirmação dos dados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   26
            ToolTipText     =   "Informe a senha do usuário autorizado"
            Top             =   1200
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmAlteraStatusFCE.frx":265E
            TabIndex        =   25
            Top             =   960
            Width           =   1935
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
            Height          =   375
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "Confirme o nº da FCE a ser CONCLUÍDA"
            Top             =   480
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmAlteraStatusFCE.frx":26D8
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Inconsistência"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1560
         Picture         =   "frmAlteraStatusFCE.frx":273E
         ScaleHeight     =   465.455
         ScaleMode       =   0  'User
         ScaleWidth      =   480
         TabIndex        =   19
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1560
         Picture         =   "frmAlteraStatusFCE.frx":3408
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1560
         Picture         =   "frmAlteraStatusFCE.frx":40D2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   12
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   480
         Left            =   1560
         Picture         =   "frmAlteraStatusFCE.frx":4D9C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   240
         Width           =   480
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Paralisada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Concluida"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Andamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da FCE "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6960
         Picture         =   "frmAlteraStatusFCE.frx":5A66
         ScaleHeight     =   465.455
         ScaleMode       =   0  'User
         ScaleWidth      =   480
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6960
         Picture         =   "frmAlteraStatusFCE.frx":6730
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6960
         Picture         =   "frmAlteraStatusFCE.frx":73FA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6960
         Picture         =   "frmAlteraStatusFCE.frx":80C4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAlteraStatusFCE.frx":8D8E
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox txtAlteraStatus 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtAlteraStatus 
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
         Height          =   375
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   6720
         OleObjectBlob   =   "frmAlteraStatusFCE.frx":8DF8
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmAlteraStatusFCE.frx":8E70
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAlteraStatusFCE.frx":8EE2
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAlteraStatusFCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 4
        If Option2.Value = True Then
            If VerificaUsuário = True Then
                'ativa rotina para ativa backup do banco
                frmAlteraStatusFCE.MousePointer = 11
                BackupZeus
                'ativa rotina que altera o status de todas as OS e das operações das OS's referentes à FCE para 3
                FechaOsFCE
                'ativa rotina de envio de e-mail
                enviaEmailFCE txtAlteraStatus(0).Text, txtAlteraStatus(1).Text, NomUsu
                AtualizaStatusFCE
                AtualizaListview
                CompoeControles
                Frame3.Visible = False
                mobjMsg.Abrir "Dados Alterados com sucesso!", Ok, informacao, "Zeus"
            'Else
            '    mobjMsg.Abrir "Senha incorreta. A FCE não pode ser CONCLUÍDA", Ok, critico, "Zeus"
            End If
        Else
            AtualizaStatusFCE
            AtualizaListview
            CompoeControles
            mobjMsg.Abrir "Dados Alterados com sucesso!", Ok, informacao, "Zeus"
        End If
    Case 5
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    CompoeControles
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    txtAlteraStatus(0).Text = varGlobal
    SkinLabel4.Caption = MeuLV.ListView1.SelectedItem.ListSubItems.Item(11)
    
    Dim rsFCE As New ADODB.Recordset
    Dim sqlFCE As String
    
    sqlFCE = "select CASE WHEN A.status = 0 THEN 'ANDAMENTO' WHEN A.status IS NULL THEN 'INCONSISTENCIA' WHEN A.status = 1 THEN 'CONCLUIDA' WHEN A.status = 2 THEN 'PARALIZADA' END AS STATUS,b.DESCRICAO from " & vBancoTotvs & ".dbo.TTB3 as b left join tbfce as a on a.fce =b.CODTB3FAT and a.fce = '" & Val(varGlobal) & "' where b.CODTB3FAT = '" & Val(varGlobal) & "'"
    rsFCE.Open sqlFCE, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFCE.RecordCount = 0 Then Exit Sub
    
    SkinLabel4.Caption = rsFCE.Fields(0)
    txtAlteraStatus(1).Text = rsFCE.Fields(1)
    
    If SkinLabel4.Caption = "ANDAMENTO" Then
        Picture4.Visible = True
        Picture5.Visible = False
        Picture6.Visible = False
        Picture8.Visible = False
        Option1.Value = True
        Option1.Enabled = False
        Option2.Enabled = True
        Option3.Enabled = True
        Option4.Enabled = False
    ElseIf SkinLabel4.Caption = "CONCLUIDA" Then
        Picture4.Visible = False
        Picture5.Visible = True
        Picture6.Visible = False
        Picture8.Visible = False
        Option2.Value = True
        Option1.Enabled = False
        Option2.Enabled = False
        Option3.Enabled = False
        Option4.Enabled = False
    ElseIf SkinLabel4.Caption = "PARALIZADA" Then
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = True
        Picture8.Visible = False
        Option3.Value = True
        Option1.Enabled = True
        Option2.Enabled = True
        Option3.Enabled = False
        Option4.Enabled = False
    ElseIf SkinLabel4.Caption = "INCONSISTENCIA" Then
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        Picture8.Visible = True
        Option4.Value = True
        Option1.Enabled = False
        Option2.Enabled = True
        Option3.Enabled = False
        Option4.Enabled = False
    End If
    rsFCE.Close
    Set rsFCE = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub AtualizaStatusFCE()
On Error GoTo Err
    Dim rsAtualizaStatus As New ADODB.Recordset
    Dim sqlAtualizaStatus As String
    
    sqlAtualizaStatus = "select * from tbfce where fce = '" & Val(txtAlteraStatus(0)) & "'"
    rsAtualizaStatus.Open sqlAtualizaStatus, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAtualizaStatus.RecordCount = 0 Then
        rsAtualizaStatus.Close
        Set rsAtualizaStatus = Nothing
        sqlAtualizaStatus = "Insert into tbfce(fce,dataabertura,cartaproposta,dataentrega) Values('" & Val(txtAlteraStatus(0)) & "','" & Format(CStr(Date), "yyyy-mm-dd") & "','" & "..." & "','" & Format(CStr(Date), "yyyy-mm-dd") & "')"
        rsAtualizaStatus.Open sqlAtualizaStatus, cnBanco
    End If
    'rsAtualizaStatus.Close
    
    If Option1.Value = True Then 'ANDAMENTO
        sqlAtualizaStatus = "Update tbfce set status = 0 where fce = '" & Val(txtAlteraStatus(0)) & "'"
    ElseIf Option2.Value = True Then 'CONCLUIDA
        sqlAtualizaStatus = "Update tbfce set status = 1 where fce = '" & Val(txtAlteraStatus(0)) & "'"
    ElseIf Option3.Value = True Then 'PARALIZADA
        sqlAtualizaStatus = "Update tbfce set status = 2 where fce = '" & Val(txtAlteraStatus(0)) & "'"
    ElseIf Option4.Value = True Then 'INCONSISTÊNCIA
        mobjMsg.Abrir "A confirmação dessa opção irá excluir os dados do cabeçalho da FCE no Zeus. Confirma?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            sqlAtualizaStatus = "Delete from tbfce where fce = '" & Val(txtAlteraStatus(0)) & "'"
        Else
            Exit Sub
        End If
    End If
    rsAtualizaStatus.Open sqlAtualizaStatus, cnBanco
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub AtualizaListview()
On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(12) = ""
    If Option1.Value = True Then 'ANDAMENTO
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(12).ReportIcon = "ANDAMENTO"
    ElseIf Option2.Value = True Then 'CONCLUIDA
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(12).ReportIcon = "CONCLUIDA"
    ElseIf Option3.Value = True Then 'PARALIZADA
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(12).ReportIcon = "PARALIZADA"
    ElseIf Option4.Value = True Then 'INCONSISTÊNCIA
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(12).ReportIcon = "DUVIDA"
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub Option1_Click()
    Frame3.Visible = False
End Sub

Private Sub Option2_Click()
    If SkinLabel4.Caption <> "CONCLUIDA" Then Frame3.Visible = True
End Sub

Private Sub Option3_Click()
    Frame3.Visible = False
End Sub

Private Sub Option4_Click()
    Frame3.Visible = False
End Sub

Private Function VerificaUsuário()
On Error GoTo Err
    VerificaUsuário = False
    
    Dim rsVerificaUsuário As New ADODB.Recordset
    Dim sqlVerificaUsuário As String
    
    If Text1.Text <> txtAlteraStatus(0).Text Then
        mobjMsg.Abrir "FCE informada não confere com a FCE que está tentando CONCLUIR.", Ok, critico, "Zeus"
        Exit Function
    End If
    
    sqlVerificaUsuário = "select b.nome from tbsenha as a inner join tbUsuarios as b on a.codigo = b.codigo where a.senha = '" & Text2.Text & "'"
    rsVerificaUsuário.Open sqlVerificaUsuário, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVerificaUsuário.RecordCount = 0 Then
        rsVerificaUsuário.Close
        Set rsVerificaUsuário = Nothing
        mobjMsg.Abrir "Senha incorreta. A FCE não pode ser CONCLUÍDA", Ok, critico, "Zeus"
        Exit Function
    End If
    If rsVerificaUsuário.Fields(0) = NomUsu Then
        VerificaUsuário = True
    End If
    rsVerificaUsuário.Close
    Set rsVerificaUsuário = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Private Sub BackupZeus()
On Error GoTo Err
    Dim rsBackRest As New ADODB.Recordset
    Dim sqlBackRest As String

    backupFile = "N'H:\usuarios\BackupZeus\BkpDADOS-" & sDatabaseName & "-" & Text1.Text & ".bak'"
    sqlBackRest = "BACKUP DATABASE [" & sDatabaseName & "] TO DISK = " & backupFile & " WITH NOFORMAT, INIT,  NAME = N'ZEUS-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
    cnBanco.CommandTimeout = 0
    rsBackRest.Open (sqlBackRest), cnBanco, adOpenStatic, adLockPessimistic
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub FechaOsFCE()
On Error GoTo Err
    Dim rsFechaOsFCE As New ADODB.Recordset
    Dim sqlFechaOsFCE As String
    
    'UPDATE NA TABELA TBMP UTILIZANDO INNER JOIN - DEU CERTO
    sqlFechaOsFCE = "UPDATE tbMP SET tbMP.status = 3 FROM tbMP INNER JOIN tbProjetos ON tbMP.codprojeto = tbProjetos.codprojeto WHERE tbProjetos.fce = '" & Val(Text1.Text) & "'"
    rsFechaOsFCE.Open sqlFechaOsFCE, cnBanco

    'UPDATE NA TABELA TBMPITENS UTILIZANDO INNER JOIN - EM TESTE
    sqlFechaOsFCE = "UPDATE tbMPItens SET tbMPItens.status = 3 FROM tbMPItens INNER JOIN tbMP ON tbMPItens.idprogramacao = tbMP.idprogramacao INNER JOIN tbProjetos ON tbMP.codprojeto = tbProjetos.codprojeto WHERE tbProjetos.fce = '" & Val(Text1.Text) & "'"
    rsFechaOsFCE.Open sqlFechaOsFCE, cnBanco
    frmAlteraStatusFCE.MousePointer = 0
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub enviaEmailFCE(vFCE As Integer, vDescricao As String, vUsuario As String)
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
'On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    vDesenhosGlobal = ""
    
    
    
    'vSMTP = "smtp.viga.ind.br"
    'vUsuEmail = "viga@viga.ind.br"
    'vSenhaEmail = "Xbkwolpb7rpd0td"
    
    'vSMTP = "smtp.viga.ind.br"
    'vUsuEmail = "taos@viga.ind.br"
    'vSenhaEmail = "taos2017@"
    
    vDecisao = "Aprovado"
    vCorDecisao = "#CD2626"

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
      
        .To = "viga@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        '.To = "viga@viga.ind.br;qualidade@viga.ind.br;almoxarifado@viga.ind.br;planejamento@viga.ind.br;orcamento@viga.ind.br;contratos@viga.ind.br;presidencia@viga.ind.br;superintendencia@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .From = "viga@viga.ind.br"  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "Conclusão da FCE nº: " & vFCE
        
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
        " <B><FONT STYLE='COLOR:#009ACD'> CONCLUSÃO DE PROJETO </FONT></B><BR><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> Foi realizado pelo usuário: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & vUsuario & "<FONT STYLE='COLOR:#009ACD'> a conclusão da FCE nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vFCE & " - " & vDescricao & "</b></FONT><BR>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & "" & " </FONT><FONT STYLE='COLOR:#009ACD'> Data da conclusão: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & Date & " </b></FONT><BR><FONT STYLE='COLOR:#009ACD'>Hora da conclusão: <FONT STYLE='COLOR:#CD2626'><b>" & Time & "</b></FONT><BR><BR><BR>" & _
        " <table border='0' cellspacing='0' width='627'>" & _
        " <tr><td width='100%'><span class='txt'>" & _
        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
        " </td></tr></table></body>"
        .Send
    End With
    'mobjMsg.Abrir "Email enviado com suscesso!", Ok, informacao, "ZEUS"
    Exit Sub
errMail:
    Msgbox "Email não enviado para o usuário solicitante do PDO." & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SGCH"
    Exit Sub
End Sub


