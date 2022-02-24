Attribute VB_Name = "RotinaGeral"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Biblioteca para manipulação do Regedit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
Public vTransacaoAtiva As Long

Public oConn As ADODB.Connection
Public sDatabaseName As String 'Utilizado para criar nova conexão com o banco na tela de splash
Public sServerName As String ' Utilizado para criar nova conexão com o banco na tela de splash
Public sUsuName As String 'Nome do usuário de conexão ao DB
Public sSenhaDB As String 'Senha de conexão ao DB
Public sSGBD As Integer 'Versão do SGBD
Public sEmailCD As String 'String que guarda endereços de e-mail que receberão notificações CD - Comunicação de Desvio
Public sEmailRNC As String 'String que guarda endereços de e-mail que receberão notificações RNC - Registro de Não Conformidade
Public sEmailSI As String 'String que guarda endereços de e-mail que receberão notificações SI - Solicitação de Inspeção
Public sEmailSRM As String 'String que guarda endereços de e-mail que receberão notificações SRM - Solicitação de Retirada de Material

Public sLogoEmpresa As String ' Utilizado para guardar o caminho da logo da empresa
Public StatusTrei As String 'Verifica o status do treinamento
Public Status As String
Public vFCE As Integer

Public rsVExp As New ADODB.Recordset
Public SqlVExp As String

Public cnBanco As ADODB.Connection
Public cnBancoTotvs As ADODB.Connection
Public rsLocal As New ADODB.Recordset
Public Sqlp As String
'Public rsResumo As New ADODB.Recordset
Public IniciaRelsEm As Integer
Public vAprovadoRest As Double
Public vAvisos As String 'Ao entrar o sistema é exibida uma tela de Avisos, onde será informado pendências no sistema
Public vCalcExp As String 'Calcula automaticamente o tempo de experiência dos colaboradores
Public GeraIntr As String 'Identifica se o sistema irá gerar ou não treinamentos introdutorios para colaboradores
Public GeraObri As String 'Identifica se o sistema irá gerar ou não treinamentos obrigatorios para colaboradores
Public GeraLog As String
Public QualForm As String

'Variaveis para armazenar dados de afastamento de colaboradores
Public vAfastDias As String
Public vAfastTreiInt As String
Public vAfastTreiObr As String
Public vStatusWin As Integer
'-------------------------------------------

Public XCodGrp As Integer 'Armazena o codigo do grupo que o usuário esta logado
Public vInc As String, vExc As String, vEdi As String, vSal As String, vImp As String, vFil As String, vAva As String, vAdi As String, vDem As String, vAdiRep As String, vAdiRes As String

Public varGlobal As String
Public varGlobal2 As String
Public FiltroGeral As String
Public Formulario As Variant
Public SqlExcLVGeral As String
Public Posicao As Integer
Public vPDO As Integer 'Variavel para armazenar ultimo numero de PDO criado
Public vCodModeloAval As Integer 'variavel do codigo do modelo de avaliação de eficacia usado na programação
Public vCodcoligada As Integer 'Variavel que armazena codigo da coligada ativa
Public vCaminhoAtu As String 'Variavel que armazena caminho + executál de atualização automática do ZEUSH

Public vControlaDim  As Integer 'Controla a quantidade de vezes q sera dimensionado o MeuLV
Public vSituacao As String 'armazena a situacao do colaborador apos a avaliacao do treinamento
Public vNota As String 'armazena a nota do colaborador/candidato apos a avaliacao do treinamento
Public vqtdava As String 'armazena a quantidade de questoes avaliadas

Public Tipo As String 'armazana o tipo de hospede (F/J) para exclusao
'Public ContaReg As Integer
Public Pesquisa As String
Public Legenda As String ' Legenda do StatusBar1
Public LegendaExc As String ' Legenda de Exclusao
Public CodUsu As String ' codigo do usuário q esta logado
Public NomUsu As String ' Nome do usuario
Public GrupoUsu As String ' Grupo do usuario
Public NomeEmpresa As String ' Nome da empresa
Public vSMTP As String 'grava endereço do servidor de SMTP
Public vUsuEmail As String 'grava nome do usuario de autenticação
Public vSenhaEmail As String 'grava a senha do usuário de autenticação
Public vIntegra As String  'Para informar se o ZEUSH esta integrado a outro sistema
Public vDataDoBanco As Date 'Grava a data atual do Banco de dados
Public vDadosTotvs(18) As String
Public colheDados(17) As String 'Guarda dados de importação de colaboradores de arquivo TXT
Public FimAprop As String 'Verifica se o colaborador tem permissão de encerra apropriação de colaboradores que estão apropriando em alguma OS

Public mStream As ADODB.Stream 'Para gravar imagem no Banco Totvs

Public vServerTotvs As String  'Armazena nome do servidor totvs
Public vBancoTotvs As String  'Armazena nome do banco totvs
Public vUsuBancoTovs As String  'Armazena usuario do banco totvs
Public vSenhaBancoTotvs As String  'Armazena senha do banco totvs

Public chamaForm As Form

Public MeuLV As New frmPesqGeral
'Public MeuLV As New frmPesqGeralTeste2
Public NomeColLV(20) As String
Public AddDadosGeral(10) As String 'Guarda dados de admissão no Processo Seletivo
Public QtdColReal As Integer 'Quantidade real de colunas do Listview
Public SqlLV As String  'Query do listview atual

Public campo As Integer
Public Campo1 As Integer
Public campo2 As Integer
Public campo3 As Integer
Public Campo4 As Integer

Public dataFilter1 As String
Public dataFilter2 As String
Public LimiteLinhas As Integer
Public Contador As Integer
Public vTime As String
Public vClone As String
Public vCodRel As Integer 'Armazena Codigo dos relatorios de inspeção
Public vTabela1 As String, vTabela2 As String, vTabela3 As String, vTabela4 As String, vTabela5 As String, vTabela6 As String, vTabela7 As String, vTabela8 As String, vTabela9 As String, vTabela10 As String

'Public CodUsu As String ' codigo do usuário q esta logado
'Public NomUsu As String ' Nome do usuario
'Public CapturaCodigo As String ' Codigo da Empresa e do Contato
'Public Legenda As String ' Informa o significado (F)Fone (F) Fax (C) Celular
Public procnom As String, procnom1 As String
Public strAno As String 'Usada no relatorio de programação de cursos/treinamentos anual
Public vQualquerDado(50, 30) As String
Public vCorTipoFCE(10, 1) As String, vGuardaLinhaTipo(5000, 1)
Public vponteiro As Integer
Public vIDCorTipoFCE As String
Public vNovoFiltro As String
Public vAlteraLike As String
Public vAlteraLike2 As String

Public Sub Main()
On Error GoTo Err1
    frmSplash.Show
    'Conexao
    'MDIPrincipal.Show
    Exit Sub
Err1:
    Msgbox "(ADO) Erro ao tentar acessar DB " & vbCrLf & Err.Number & " - Procure o administrador da rede " & Err.Description, 16, "Mensagem de erro"
    'mobjMsg.Abrir "(ADO) Erro ao tentar acessar DB " & vbCrLf & Err.Number & " - Procure o administrador da rede " & Err.Description, 16, "Mensagem de erro", ok, critico, "Atenção"
    Exit Sub
End Sub

Public Function Conexao()
On Error GoTo Err
    Conexao = True
    If sServerName = "" Then GoTo Err
    Set cnBanco = New ADODB.Connection
    'If sSGBD = 1 Then
    '    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=False;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    'ElseIf sSGBD = 2 Then
        cnBanco.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=True;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    'Else
    '    Resume Err1
    'End If
    frmSplash.Label5.Caption = "Conexão realizada com sucesso"
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    ElseIf Err.Number = 3705 Then
        cnBanco.Close
        Resume
    Else
        Conexao = False
        frmSplash.Label5.Caption = "Falha na conexão"
        Exit Function
    End If
End Function

'ABAIXO CONEXÃO COM O BANCO DE DADOS RM
Public Function ConexaoTotvs()
On Error GoTo Err
    ConexaoTotvs = True
    Set cnBancoTotvs = New ADODB.Connection
    cnBancoTotvs.Open "Provider=SQLOLEDB.1;Password=" & vSenhaBancoTotvs & ";Persist Security Info=True;User ID=" & vUsuBancoTovs & ";Initial Catalog=" & vBancoTotvs & ";Data Source=" & vServerTotvs
    vIntegra = "S"
    'achaSecaoZEUSH
    'criaTrigger
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    ElseIf Err.Number = 3705 Then
        cnBancoTotvs.Close
        Resume
    Else
        ConexaoTotvs = False
        mobjMsg.Abrir "Erro de conexão com Banco Totvs", Ok, critico, "Atenção"
        vIntegra = "N"
        Exit Function
    End If
End Function

Public Function criaTrigger()
On Error GoTo Err
    'Essa rotina cria uma TRIGGER entre o Banco da TOTVS e o Banco do ZEUSH
    Dim sqlTrigger As String
    Dim rsTrigger As New ADODB.Recordset
    Dim rsVerTabZEUSH As ADODB.Recordset
    
'Verifica se a tabela zZEUSH_Demitidos existe no banco CORPORERM
    Set rsVerTabZEUSH = cnBancoTotvs.OpenSchema(adSchemaTables, Array(Empty, Empty, "zZEUSH_Demitidos", "Table"))
    If rsVerTabZEUSH.EOF Then
        cnBancoTotvs.Execute "CREATE TABLE " & vBancoTotvs & ".dbo.zZEUSH_Demitidos(" & _
        "chapa VARCHAR(10) NOT NULL," & _
        "controleZEUSH CHAR(1) NOT NULL," & _
        "PRIMARY KEY (chapa))"
    End If
    rsVerTabZEUSH.Close
'FIM TESTE
    
    sqlTrigger = "CREATE TRIGGER TriggerMonitoraTotvs on PFUNC for insert,update as Insert dbo.zZEUSH_Demitidos " & _
                "Select CHAPA,'' from inserted Where CODSITUACAO = 'D'"
    
    'sqlTrigger = "CREATE TRIGGER [dbo].[TriggerMonitoraTotvs] on [dbo].[PFUNC]For insert,update as if (select count (*) from deleted) <> 0 " & _
    '            "Update B set B.ativo = 'N',B.datarecisao = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))),B.homologacaonum = 'Ver Totvs',b.homologacaoorgao = 'Ver Totvs'FROM dbo.PFUNC as A Inner join " & _
    '            sDatabaseName & ".dbo.tbcolaboradores as B on A.CHAPA=B.CODCOLABORADOR COLLATE SQL_Latin1_General_CP1_CI_AS Where A.CODSITUACAO = 'D'"
    rsTrigger.Open sqlTrigger, cnBancoTotvs
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        sqlTrigger = "ALTER TRIGGER [dbo].[TriggerMonitoraTotvs] on [dbo].[PFUNC] for insert,update as Insert dbo.zZEUSH_Demitidos " & _
                    "Select CHAPA,'' from inserted Where CODSITUACAO = 'D'"
        rsTrigger.Open sqlTrigger, cnBancoTotvs
    End If
End Function

Public Sub achaSecaoZEUSH()
On Error GoTo Err
    Dim sqlSecaoZEUSH As String
    Dim rsSecaoZEUSH As New ADODB.Recordset
    Dim vIDSecao As Integer
    
    Dim sqlGravaZEUSH As String
    Dim rsGravaZEUSH As New ADODB.Recordset
    
    
    'sqlSecaoZEUSH = "Select MAX(id)+1 from PSECAO"
    'rsSecaoZEUSH.Open sqlSecaoZEUSH, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    'vIDSecao = rsSecaoZEUSH.Fields(0)
    'rsSecaoZEUSH.Close
    'Set rsSecaoZEUSH = Nothing
    
    'Cria seção de ADMISSÃO de colaboradores do ZEUSH
    sqlSecaoZEUSH = "select * from PSECAO where DESCRICAO = 'ZEUSH'"
    rsSecaoZEUSH.Open sqlSecaoZEUSH, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    If rsSecaoZEUSH.EOF Then
        sqlGravaZEUSH = "Insert into " & _
                       "PSECAO(codcoligada,codigo,descricao,cgc,fpas,sat,rua,numero,bairro,estado,cidade,cep,pais,telefone,naoempregpropr,categoria,codterceirosinss," & _
                                "PERCENTTERCEIROS,percentacidtrab,proprantes5dia1,proprantes5dia2,centrantes5dia1,centrantes5dia2,CONTRIBSESIESENAI,distribpetroleo,pessoafisica," & _
                                "secaodesativada,identificacaocgc,enderecoalterou,codmunicipio,naturezajuridica,codcalendario,prefixoinscrfgts,primeiradeclcaged,encerramento," & _
                                "codfilial,coddepto,optasimples,alteracaocaged,codpagtogps,participapat,porteempresa,ddd,isentocontribsocial,vincpat5sal,vincpatmaior5sal,porcservprop," & _
                                "porcadmcozinha,porcrefeicaoconv,porcrefeicaotransp,porccestaalimento,PORCALIMCONVENIO,email,cnaerais,valorentidadesacumulado,idmemoambtrab,visivelorganograma," & _
                                "codigopai,reccreatedby,reccreatedon,recmodifiedby,recmodifiedon) " & _
                       "Values(1,'001.01.01.01','ZEUSH','19.431.980/0001-05','507','2511000','AV VITO GAGGIATO','SN','DISTRITO INDUSTRIAL','MG','SANTANA DO PARAISO', " & _
                                "'35167-000','BRASIL','3801-2600',3,'99','0079',5.80,3.00,0,0,0,0,0,0,0,0,1,0,'3158953','2062','0000001','01',1,2,1,'01',1,2,'2100', " & _
                                "0,2,'0031', 0,0,0,0,0,0,0,0,0,'pessoal@viga.ind.br','25110','0.0000',82,'T','001.01','mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ",'mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ")"
        rsGravaZEUSH.Open sqlGravaZEUSH, cnBancoTotvs
    End If
    rsSecaoZEUSH.Close
    Set rsSecaoZEUSH = Nothing
    
    'Cria seção de ALTERÇÃO FUNCIONAL de colaboradores do ZEUSH
    sqlSecaoZEUSH = "select * from PSECAO where DESCRICAO = 'ZEUSH - Alteração funcional'"
    rsSecaoZEUSH.Open sqlSecaoZEUSH, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    If rsSecaoZEUSH.EOF Then
        sqlGravaZEUSH = "Insert into " & _
                       "PSECAO(codcoligada,codigo,descricao,cgc,fpas,sat,rua,numero,bairro,estado,cidade,cep,pais,telefone,naoempregpropr,categoria,codterceirosinss," & _
                                "PERCENTTERCEIROS,percentacidtrab,proprantes5dia1,proprantes5dia2,centrantes5dia1,centrantes5dia2,CONTRIBSESIESENAI,distribpetroleo,pessoafisica," & _
                                "secaodesativada,identificacaocgc,enderecoalterou,codmunicipio,naturezajuridica,codcalendario,prefixoinscrfgts,primeiradeclcaged,encerramento," & _
                                "codfilial,coddepto,optasimples,alteracaocaged,codpagtogps,participapat,porteempresa,ddd,isentocontribsocial,vincpat5sal,vincpatmaior5sal,porcservprop," & _
                                "porcadmcozinha,porcrefeicaoconv,porcrefeicaotransp,porccestaalimento,PORCALIMCONVENIO,email,cnaerais,valorentidadesacumulado,idmemoambtrab,visivelorganograma," & _
                                "codigopai,reccreatedby,reccreatedon,recmodifiedby,recmodifiedon) " & _
                       "Values(1,'001.01.01.02','ZEUSH - Alteração funcional','19.431.980/0001-05','507','2511000','AV VITO GAGGIATO','SN','DISTRITO INDUSTRIAL','MG','SANTANA DO PARAISO', " & _
                                "'35167-000','BRASIL','3801-2600',3,'99','0079',5.80,3.00,0,0,0,0,0,0,0,0,1,0,'3158953','2062','0000001','01',1,2,1,'01',1,2,'2100', " & _
                                "0,2,'0031', 0,0,0,0,0,0,0,0,0,'pessoal@viga.ind.br','25110','0.0000',82,'T','001.01','mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ",'mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ")"
        rsGravaZEUSH.Open sqlGravaZEUSH, cnBancoTotvs
    End If
    rsSecaoZEUSH.Close
    Set rsSecaoZEUSH = Nothing
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

Public Sub localizaCorTipoFCE(nomeTipoFCE As String)
    Dim X As Integer
    vIDCorTipoFCE = ""
    For X = 0 To 10
        If nomeTipoFCE = vCorTipoFCE(X, 0) Then
            vIDCorTipoFCE = vCorTipoFCE(X, 1)
        End If
        If vCorTipoFCE(X, 0) = "" Then Exit For
    Next
End Sub

Public Sub CompoeCombo(Combo As ComboBox, Tabela, campo, Campo1)
On Error GoTo Err
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    'Se a tabela for tbsetores, somente irá exibir os setores ativos
    If Tabela = "tbsetores" Then
        sql = "Select * from " & Tabela & " where codcoligada = '" & vCodcoligada & "' and ativo = 'S' Order By " & Campo1
    Else
        sql = "Select * from " & Tabela & " where codcoligada = '" & vCodcoligada & "' Order By " & Campo1
    End If
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsTabela.EOF() Then
'        Combo.Clear
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            Combo.AddItem rsTabela.Fields(Campo1)
            Combo.ItemData(Combo.NewIndex) = Val(rsTabela.Fields(0))
            rsTabela.MoveNext
        Next
    End If
    rsTabela.Close
    Set rsTabela = Nothing
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

Public Sub CompoeCombo1(Combo As ComboBox, Tabela, campo, Campo1)
On Error GoTo Err
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    sql = "Select * from " & Tabela & " where codcoligada = '" & vCodcoligada & "' Order By " & campo
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsTabela.EOF() Then
'        Combo.Clear
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            Combo.AddItem Format(rsTabela.Fields(campo), "000000") & "-" & rsTabela.Fields(Campo1)
            Combo.ItemData(Combo.NewIndex) = Val(rsTabela.Fields(0))
            rsTabela.MoveNext
        Next
    End If
    rsTabela.Close
    Set rsTabela = Nothing
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

Public Sub CompoeCombo2(Combo As ComboBox, Tabela, campo, Campo1)
On Error GoTo Err
    'Idependente de coligada
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    'Se a tabela for tbsetores, somente irá exibir os setores ativos
    If Tabela = "tbsetores" Then
        sql = "Select * from " & Tabela & " where ativo = 'S' Order By " & Campo1
    ElseIf Tabela = "tbDesConjunto" Then
        'Nesse caso CAMPO recebe o nome da FCE
        sql = "select a.idConjunto from tbDesConjunto as a inner join tbDesenhos as b on a.iddesenho = b.iddesenho inner join tbProjetos as c on b.codprojeto = c.codprojeto where c.fce = '" & campo & "' group by a.idConjunto"
        Combo.AddItem "-"
    Else
        sql = "Select " & Campo1 & "," & campo & " from " & Tabela & " Order By " & Campo1
    End If
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsTabela.EOF() Then
'        Combo.Clear
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            Combo.AddItem rsTabela.Fields(0)
            Combo.ItemData(Combo.NewIndex) = Val(rsTabela.Fields(1))
            rsTabela.MoveNext
        Next
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    Exit Sub
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
        Exit Sub
    End If
End Sub

Public Sub CompoeCampoCombo(Codigo, Combo As ComboBox)
    Dim X As Integer
    For X = 0 To Combo.ListCount - 1
        Combo.ListIndex = X
        If Combo.List(X) = Codigo Then
            Exit For
        End If
    Next
End Sub

Public Sub CompoeComboLV(Combo As ComboBox, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    Combo.Clear
    'If Column Is Nothing Then
        For Each c In MeuLV.ListView1.ColumnHeaders
            Combo.AddItem c
        Next
        Combo.Text = Combo.List(0)
    'End If
End Sub

Public Sub CompoeComboLVPesq(Combo As ComboBox, LV As Listview, vIndiceCombo As Integer, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    'If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            Combo.AddItem c
        Next
        Combo.Text = Combo.List(vIndiceCombo)
    'End If
End Sub

Public Sub CompoeComboCC(Combo As ComboBox)
On Error GoTo Err
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    sql = "select a.NOME from " & vBancoTotvs & ".dbo.GCCUSTO as a where a.ATIVO = 'T' and substring(a.nome,1,4) = '3000' or substring(a.nome,1,4) = '4000' or substring(a.nome,1,4) = '7000' or substring(a.nome,1,4) = '5000'"
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
    Combo.Clear
    If Not rsTabela.EOF() Then
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            Combo.AddItem rsTabela.Fields(0)
            rsTabela.MoveNext
        Next
    Else
        Combo.AddItem ("-")
        Combo.Text = "-"
    End If
    rsTabela.Close
    Set rsTabela = Nothing
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

Public Sub CompoeComboSQL(Combo As ComboBox, vSql As String)
On Error GoTo Err
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    rsTabela.Open vSql, cnBanco, adOpenKeyset, adLockReadOnly
    Combo.Clear
    If Not rsTabela.EOF() Then
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            Combo.AddItem rsTabela.Fields(0)
            rsTabela.MoveNext
        Next
    Else
        Combo.AddItem ("-")
        Combo.Text = "-"
    End If
    rsTabela.Close
    Set rsTabela = Nothing
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

Public Function RemoveMask(campo)
    Dim Variavel As String
    Dim Varival As String
    Variavel = Replace(campo, ":", "")
    RemoveMask = Variavel
End Function

Public Function RemoveMask2(campo, vChar)
    Dim Variavel As String
    Dim Varival As String
    Variavel = Replace(campo, vChar, "")
    RemoveMask2 = Variavel
End Function

Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function

Public Function CriarTabelasADO() As Boolean
On Error GoTo Err
    
    Dim rsSenha As New ADODB.Recordset
    Set rsSenha = New ADODB.Recordset
    
    Dim rsUsuario As New ADODB.Recordset
    Set rsUsuario = New ADODB.Recordset
    
    Dim rsGrupo As New ADODB.Recordset
    Set rsGrupo = New ADODB.Recordset
    
    Dim rsConfGrupo As New ADODB.Recordset
    Set rsConfGrupo = New ADODB.Recordset
    
    Dim SqlSenha As String
    Dim SqlUsuario As String
    Dim SqlGrupo As String
    Dim SqlConfGrupo As String
    Dim Y As Integer, X As Integer
    
    sServerName = frmSplash.Combo1.Text
    sDatabaseName = frmSplash.Combo2.Text
    Set oConn = New ADODB.Connection
    
    'oConn.Open "Provider=SQLOLEDB;Data Source=" & sServerName & ";User ID=sa;Password=;"
    oConn.Open "Provider=SQLOLEDB;Data Source=" & sServerName & ";User ID=" & sUsuName & ";Password=" & sSenhaDB & ";"
    
    'CRIA BANCO ZEUSH
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDadosBanco(" & _
    "NomeServidor VARCHAR(50) NULL," & _
    "NomeBanco VARCHAR(50) NULL)"
    
    'TABELAS ZEUS
'============================
    'CRIA TABELAS ZEUS
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTipoMat(" & _
    "codigo NUMERIC NOT NULL," & _
    "descricao VARCHAR(50) NOT NULL," & _
    "ativo VARCHAR(1) NOT NULL," & _
    "PRIMARY KEY (codigo))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAtividades(" & _
    "codigo NUMERIC NOT NULL," & _
    "descricao VARCHAR(100) NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "PRIMARY KEY (codigo))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbVerifGrupo(" & _
    "codgrupo NUMERIC NOT NULL," & _
    "descricao VARCHAR(100) NOT NULL," & _
    "PRIMARY KEY (codgrupo))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbVerifItem(" & _
    "codgrupo NUMERIC NOT NULL," & _
    "coditem NUMERIC NOT NULL," & _
    "descricao VARCHAR(100) NOT NULL," & _
    "PRIMARY KEY (codgrupo,coditem))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbListaVerif(" & _
    "fce NUMERIC NOT NULL," & _
    "codgrupo NUMERIC NOT NULL," & _
    "coditem NUMERIC NOT NULL," & _
    "observacao VARCHAR(100) NULL," & _
    "PRIMARY KEY (fce,codgrupo,coditem))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFCE(" & _
    "fce NUMERIC NOT NULL," & _
    "dataabertura DATETIME NOT NULL," & _
    "cartaproposta VARCHAR(50) NOT NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "obscomercial VARCHAR(300) NULL," & _
    "obsfinanceira VARCHAR(300) NULL," & _
    "dataentrega DATETIME NOT NULL," & _
    "fabricacao VARCHAR(100) NULL," & _
    "reparo VARCHAR(100) NULL," & _
    "materiaprima VARCHAR(100) NULL," & _
    "transporte VARCHAR(100) NULL," & _
    "pintura VARCHAR(100) NULL," & _
    "PRIMARY KEY (fce))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbPedidos(" & _
    "fce NUMERIC NOT NULL," & _
    "numoc NUMERIC NOT NULL," & _
    "descricao VARCHAR(100) NOT NULL," & _
    "quantidade FLOAT NOT NULL," & _
    "unqtd VARCHAR(5) NOT NULL," & _
    "peso FLOAT NOT NULL," & _
    "unpeso VARCHAR(5) NOT NULL," & _
    "valorcimp FLOAT NOT NULL," & _
    "pis FLOAT NOT NULL," & _
    "cofins FLOAT NOT NULL," & _
    "icms FLOAT NOT NULL," & _
    "ipi FLOAT NOT NULL," & _
    "bcalcicms NUMERIC NOT NULL," & _
    "tipofcedesc VARCHAR(50) NULL," & _
    "tiporceid INT NULL," & _
    "PRIMARY KEY (fce,numoc))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMateriais(" & _
    "IDPRD INT NOT NULL," & _
    "formula VARCHAR(200) NOT NULL," & _
    "constpint FLOAT NOT NULL," & _
    "forpint VARCHAR(200) NULL," & _
    "observacao TEXT NULL," & _
    "PRIMARY KEY (IDPRD))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConstantes(" & _
    "IDPRD INT NOT NULL," & _
    "valconst FLOAT NOT NULL," & _
    "descricao VARCHAR(50) NOT NULL," & _
    "idseq INT NOT NULL," & _
    "PRIMARY KEY (IDPRD, idseq))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbListaMaterial(" & _
    "codfo NUMERIC NOT NULL," & _
    "codseq NUMERIC NOT NULL," & _
    "desenho VARCHAR(50) NULL," & _
    "codmat NUMERIC NOT NULL," & _
    "quantcj NUMERIC NOT NULL," & _
    "dimensoes VARCHAR(50) NOT NULL," & _
    "pesounit FLOAT NOT NULL," & _
    "area FLOAT NOT NULL," & _
    "TipoMat NUMERIC NULL," & _
    "revisao VARCHAR(2) NULL," & _
    "observacao VARCHAR(100) NULL," & _
    "PRIMARY KEY (codfo, codseq))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFo(" & _
    "codfo NUMERIC NOT NULL," & _
    "datafo DATETIME NOT NULL," & _
    "statusfo NUMERIC NOT NULL," & _
    "fce NUMERIC NULL," & _
    "pedido VARCHAR(50) NULL," & _
    "codclifor NUMERIC NOT NULL," & _
    "codcontato NUMERIC NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "descricao VARCHAR(100) NULL," & _
    "datadevcp DATETIME NULL," & _
    "proposta VARCHAR(50) NULL," & _
    "quantidade FLOAT NULL," & _
    "unidade VARCHAR(5) NULL," & _
    "valorunit FLOAT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "PRIMARY KEY (codfo))"
       
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbresumo(" & _
    "codfo NUMERIC NOT NULL," & _
    "codmat NUMERIC NOT NULL," & _
    "dimensoes VARCHAR(50) NOT NULL," & _
    "pesomp FLOAT NOT NULL," & _
    "aliquota NUMERIC NOT NULL," & _
    "quantidade NUMERIC NOT NULL," & _
    "peso FLOAT NULL," & _
    "TipoMat NUMERIC NULL," & _
    "codres NUMERIC NOT NULL," & _
    "observacao VARCHAR(100) NULL," & _
    "PRIMARY KEY (codfo,codmat,codres))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbresumolm(" & _
    "fce NUMERIC NOT NULL," & _
    "codlm NUMERIC NOT NULL," & _
    "codmat NUMERIC NULL," & _
    "tipomat NUMERIC NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "codres NUMERIC NOT NULL," & _
    "PRIMARY KEY (fce,codlm,codres))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFaturamento(" & _
    "fce NUMERIC NOT NULL," & _
    "numnota VARCHAR(20) NOT NULL," & _
    "data DATETIME NULL," & _
    "quantidade FLOAT NULL," & _
    "unidade VARCHAR(5) NULL," & _
    "valor FLOAT NULL," & _
    "PRIMARY KEY (fce,numnota))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbLM(" & _
    "fce NUMERIC NOT NULL," & _
    "codlm NUMERIC NOT NULL," & _
    "dataabertura DATETIME NOT NULL," & _
    "descricao VARCHAR(100) NULL," & _
    "ld VARCHAR(100) NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "tipofceid INT NULL," & _
    "tipofcedesc VARCHAR(50) NULL," & _
    "PRIMARY KEY (fce,codlm))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbItemLM(" & _
    "fce NUMERIC NOT NULL," & _
    "codlm NUMERIC NOT NULL," & _
    "codseq NUMERIC NOT NULL," & _
    "codigodes NUMERIC NULL," & _
    "codigopos NUMERIC NULL," & _
    "codmat INT NULL," & _
    "quantcj NUMERIC NULL," & _
    "quantunit NUMERIC NULL," & _
    "dimensoes VARCHAR(50) NULL," & _
    "pesounit FLOAT NOT NULL," & _
    "area FLOAT NOT NULL," & _
    "tipomat NUMERIC NULL," & _
    "codfo NUMERIC NULL," & _
    "observação VARCHAR(300) NULL," & _
    "matncadast VARCHAR(300) NULL," & _
    "calcpor VARCHAR(20) NULL," & _
    "idconjunto INT NULL," & _
    "PRIMARY KEY (fce, codlm, codseq))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbLD(" & _
    "fce NUMERIC NOT NULL," & _
    "codprojeto NUMERIC NOT NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "codresponsavel NUMERIC NOT NULL," & _
    "status NUMERIC NULL," & _
    "PRIMARY KEY (fce,codprojeto))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbItemLD(" & _
    "fce NUMERIC NOT NULL," & _
    "codprojeto NUMERIC NOT NULL," & _
    "idld NUMERIC NOT NULL," & _
    "coddesenho NUMERIC NOT NULL," & _
    "codposicao NUMERIC NOT NULL," & _
    "revisao VARCHAR(2) NOT NULL," & _
    "pesolm FLOAT NOT NULL," & _
    "pesold FLOAT NOT NULL," & _
    "quantidade NUMERIC NOT NULL," & _
    "observacao VARCHAR(200) NULL," & _
    "datacad DATETIME NOT NULL," & _
    "codprocesso NUMERIC NULL," & _
    "status VARCHAR(20) NULL," & _
    "unidade VARCHAR(3) NULL," & _
    "PRIMARY KEY (fce, codprojeto, idld))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDesenhos(" & _
    "iddesenho INT NOT NULL," & _
    "datacadastro DATETIME NOT NULL," & _
    "codprojeto INT NOT NULL," & _
    "desenho VARCHAR(50) NOT NULL," & _
    "revisao VARCHAR(5) NOT NULL," & _
    "descricao TEXT NULL," & _
    "tipo VARCHAR(15) NOT NULL," & _
    "ativo CHAR(1) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (iddesenho,codcoligada))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDesenhosrev(" & _
    "iddesenho INT NOT NULL," & _
    "revisao VARCHAR(10) NOT NULL," & _
    "data DATETIME NOT NULL," & _
    "detalhes TEXT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (iddesenho,revisao,codcoligada))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDesenho(" & _
    "codigodes NUMERIC NOT NULL," & _
    "desenho VARCHAR(50) NOT NULL," & _
    "revisao VARCHAR(2) NOT NULL," & _
    "descricaodesenho VARCHAR(100) NULL," & _
    "status VARCHAR(100) NULL," & _
    "PRIMARY KEY (codigodes))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbPosicoes(" & _
    "codigopos NUMERIC NOT NULL," & _
    "codigodes NUMERIC NULL," & _
    "posicao VARCHAR(50) NULL," & _
    "descposicao VARCHAR(300) NULL," & _
    "item VARCHAR(50) NULL," & _
    "rastro VARCHAR(50) NULL," & _
    "codigoos NUMERIC NULL," & _
    "PRIMARY KEY (codigopos))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDesConjunto(" & _
    "codlm INT NOT NULL," & _
    "idseq INT NOT NULL," & _
    "iddesenho INT NOT NULL," & _
    "quantidade NUMERIC NOT NULL," & _
    "posicao VARCHAR(100) NOT NULL," & _
    "PRIMARY KEY (codlm,iddesenho))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProjetos(" & _
    "codprojeto NUMERIC NOT NULL," & _
    "fce NUMERIC NULL," & _
    "projeto VARCHAR(50) NULL," & _
    "descricao VARCHAR(300) NULL," & _
    "data DATETIME NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "oc VARCHAR(50) NULL," & _
    "PRIMARY KEY (codprojeto))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProcessos(" & _
    "codprocesso NUMERIC NOT NULL," & _
    "descricao VARCHAR(100) NOT NULL," & _
    "PRIMARY KEY (codprocesso))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFases(" & _
    "codprocesso NUMERIC NOT NULL," & _
    "codfase NUMERIC NOT NULL," & _
    "descricao VARCHAR(100) NOT NULL," & _
    "relger VARCHAR(1) NULL," & _
    "pesofab NUMERIC NULL," & _
    "titulofase VARCHAR(100) NULL," & _
    "PRIMARY KEY (codprocesso, codfase))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbos(" & _
    "idos NUMERIC NOT NULL," & _
    "rastreabilidade DATETIME NOT NULL," & _
    "observacao DATETIME NOT NULL," & _
    "dataos DATETIME NOT NULL," & _
    "revisao VARCHAR(5) NULL," & _
    "status VARCHAR(10) NULL," & _
    "tipoos INT NULL," & _
    "PRIMARY KEY (idos))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbApropriacao(" & _
    "idaprop NUMERIC NOT NULL," & _
    "codigoos NUMERIC NOT NULL," & _
    "fase varchar(30) NOT NULL," & _
    "codigofuncionario NUMERIC NOT NULL," & _
    "dataini DATETIME NOT NULL," & _
    "horaini DATETIME NOT NULL," & _
    "datafim DATETIME NULL," & _
    "horafim DATETIME NULL," & _
    "PRIMARY KEY (idaprop,codigoos))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbRelatorios(" & _
    "codrel NUMERIC NOT NULL," & _
    "fce NUMERIC NOT NULL," & _
    "codprojeto NUMERIC NOT NULL," & _
    "datarel DATETIME NOT NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "statusimp VARCHAR(1) NULL," & _
    "norma VARCHAR(20) NULL," & _
    "tiporel NUMERIC NOT NULL," & _
    "pesobalanca FLOAT NULL," & _
    "PRIMARY KEY (codrel))"
    'tiporel é referente ao tipo de relatorio: 0 = Inspeção e 1 = Expedição
    'statusimp é referente à impressão
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbItemRelatorio(" & _
    "codrel NUMERIC NOT NULL," & _
    "idld NUMERIC NOT NULL," & _
    "codprocesso NUMERIC NOT NULL," & _
    "codfase NUMERIC NOT NULL," & _
    "qtdlib FLOAT NOT NULL," & _
    "PRIMARY KEY (codrel,idld))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTranspRel(" & _
    "codrel NUMERIC NOT NULL," & _
    "codtransp NUMERIC NOT NULL," & _
    "placacavalo VARCHAR(10) NOT NULL," & _
    "ufcavalo VARCHAR(2) NOT NULL," & _
    "placacarreta VARCHAR(10) NOT NULL," & _
    "ufcarreta VARCHAR(2) NOT NULL," & _
    "nomemotorista VARCHAR(50) NULL," & _
    "PRIMARY KEY (codrel))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTransportadoras(" & _
    "codtransp NUMERIC NOT NULL," & _
    "nome VARCHAR(100) NOT NULL," & _
    "cnpj VARCHAR(30) NULL," & _
    "ie VARCHAR(30) NULL," & _
    "endereco VARCHAR(50) NULL," & _
    "cep VARCHAR(15) NULL," & _
    "bairro VARCHAR(50) NULL," & _
    "cidade VARCHAR(50) NULL," & _
    "uf VARCHAR(2) NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "PRIMARY KEY (codtransp))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfiguracoes(" & _
    "codtransp NUMERIC NOT NULL," & _
    "nome VARCHAR(100) NOT NULL," & _
    "cnpj VARCHAR(30) NULL," & _
    "ie VARCHAR(30) NULL," & _
    "endereco VARCHAR(50) NULL," & _
    "cep VARCHAR(15) NULL," & _
    "bairro VARCHAR(50) NULL," & _
    "cidade VARCHAR(50) NULL," & _
    "uf VARCHAR(2) NULL," & _
    "PRIMARY KEY (codtransp))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMenuConf(" & _
    "idmenu NUMERIC NOT NULL," & _
    "idsub VARCHAR(10) NOT NULL," & _
    "tipo VARCHAR(20) NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "id INT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "icon INT NOT NULL," & _
    "PRIMARY KEY (idmenu,idsub,tipo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbclifor(" & _
    "codclifor NUMERIC NOT NULL," & _
    "endereco VARCHAR(120) NULL," & _
    "cep VARCHAR(50) NULL," & _
    "bairro VARCHAR(50) NULL," & _
    "cidade VARCHAR(50) NULL," & _
    "uf VARCHAR(2) NULL," & _
    "telefone VARCHAR(15) NULL," & _
    "fax VARCHAR(15) NULL," & _
    "email VARCHAR(100) NULL," & _
    "site VARCHAR(50) NULL," & _
    "tipo NUMERIC NULL," & _
    "especificacao NUMERIC NULL," & _
    "codatividade NUMERIC NULL," & _
    "nome VARCHAR(100) NULL," & _
    "tipolig VARCHAR(15) NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "PRIMARY KEY (codclifor))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbContatos(" & _
    "codclifor NUMERIC NOT NULL," & _
    "codcontato NUMERIC NOT NULL," & _
    "nome VARCHAR(100) NOT NULL," & _
    "departamento VARCHAR(50) NULL," & _
    "cargo VARCHAR(50) NULL," & _
    "funcao VARCHAR(50) NULL," & _
    "telefone VARCHAR(15) NULL," & _
    "fax VARCHAR(15) NULL," & _
    "celular VARCHAR(15) NULL," & _
    "email VARCHAR(100) NULL," & _
    "ramal VARCHAR(10) NULL," & _
    "tipolig VARCHAR(15) NULL," & _
    "PRIMARY KEY (codclifor, codcontato))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFisica(" & _
    "codclifor NUMERIC NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "identidade VARCHAR(50) NULL," & _
    "cpf VARCHAR(50) NULL," & _
    "PRIMARY KEY (codclifor))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbJuridica(" & _
    "codclifor NUMERIC NOT NULL," & _
    "razsocial VARCHAR(100) NULL," & _
    "nomefantasia VARCHAR(100) NULL," & _
    "cnpj VARCHAR(50) NULL," & _
    "inscest VARCHAR(50) NULL," & _
    "PRIMARY KEY (codclifor))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbCD(" & _
    "idcd NUMERIC NOT NULL," & _
    "fce VARCHAR(50) NOT NULL," & _
    "desenho VARCHAR(50) NOT NULL," & _
    "revisao VARCHAR(5) NOT NULL," & _
    "quantidade NUMERIC NOT NULL," & _
    "pesounit FLOAT NOT NULL," & _
    "datarecebido DATETIME NULL," & _
    "ptempo VARCHAR(5) NOT NULL," & _
    "punidade VARCHAR(10) NOT NULL," & _
    "usuario VARCHAR(100) NOT NULL," & _
    "dataini DATETIME NULL," & _
    "datafim DATETIME NULL," & _
    "croqui VARCHAR(50) NULL," & _
    "status INT NOT NULL," & _
    "observacao TEXT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "detalhista VARCHAR(100) NULL," & _
    "iddesenho INT NOT NULL," & _
    "PRIMARY KEY (idcd))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.TBPONTO(" & _
    "ID INT NOT NULL IDENTITY," & _
    "DATABATIDA DATE NOT NULL," & _
    "CHAPA VARCHAR(10) NOT NULL," & _
    "BATIDA1 TIME(5) NOT NULL," & _
    "BATIDA2 TIME(5) NULL," & _
    "BATDA3 TIME(5) NULL," & _
    "BATIDA4 TIME(5) NULL," & _
    "BATIDA5 TIME(5) NULL," & _
    "BATIDA6 TIME(5) NULL," & _
    "CONTBATIDA INT NOT NULL," & _
    "PRIMARY KEY (ID))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.TBHORARIOS(" & _
    "ID INT NOT NULL IDENTITY," & _
    "CHAPA VARCHAR(10) NOT NULL," & _
    "HORARIO_ENTRADA DATETIME NOT NULL," & _
    "HORARIO_SAIDA DATETIME NOT NULL," & _
    "CODCOLIGADA INT NOT NULL," & _
    "PRIMARY KEY (ID))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.TBCONTROLEATIVIDADES(" & _
    "ID INT NOT NULL IDENTITY," & _
    "DESCRICAO TEXT NOT NULL," & _
    "DATAHORA DATETIME NOT NULL," & _
    "PRIMARY KEY (ID))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.TBCONFRELOGIO(" & _
    "IDRELOGIO VARCHAR(20) NOT NULL, " & _
    "IPRELOGIO VARCHAR(20) NOT NULL, " & _
    "CPFRESPONSAVEL VARCHAR(20) NOT NULL, " & _
    "PASSWORD DATETIME NOT NULL, " & _
    "CAMINHO VARCHAR(300) NOT NULL, " & _
    "PRIMARY KEY (IDRELOGIO))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.TBCONFFLEXJR(" & _
    "USUARIO VARCHAR(20) NOT NULL," & _
    "PASSWORD VARCHAR(20) NOT NULL, " & _
    "CAMINHO VARCHAR(300) NOT NULL)"
    
    
    'ZEUS - CRIA FUNÇÃO QUE CONVERTE MINUTOS EM HORAS:MINUTOS
    oConn.Execute "CREATE FUNCTION dbo.FN_CONVMIN(@MINUTOS int) " & _
    "RETURNS NVARCHAR(7)" & _
    "BEGIN" & _
    "DECLARE @iHoras INTEGER" & _
    "DECLARE @iMinutos INTEGER" & _
    "DECLARE @sEdita VARCHAR(7)" & _
    "SET @iHoras = CAST(ROUND(@MINUTOS/60, 0) AS INT)" & _
    "SET @iMinutos = @MINUTOS % 60" & _
    "SET @sEdita = " & _
    "CASE LEN(@iHoras)" & _
    "    WHEN 0 THEN '00'" & _
    "    WHEN 1 THEN '0' + CONVERT(NVARCHAR(1), @iHoras)" & _
    "    ELSE CONVERT(NVARCHAR(4),@iHoras)" & _
    "END" & _
    "SET @sEdita = " & _
    "@sEdita + ':' + CASE LEN(@iMinutos)" & _
    "    WHEN 0 THEN '00'" & _
    "    WHEN 1 THEN '0' + CONVERT(NVARCHAR(3), @iMinutos)" & _
    "    ELSE CONVERT(NVARCHAR(4), @iMinutos)" & _
    "END" & _
    "IF @sEdita = '00:00' BEGIN SET @sEdita = ' ' END" & _
    "RETURN @sEdita" & _
    "END GO"
    
    'CRIA TABELAS ADMINISTRATIVAS
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFuncionarios(" & _
    "codigo NUMERIC NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "setor VARCHAR(50) NOT NULL," & _
    "função VARCHAR(50) NOT NULL," & _
    "salario FLOAT NULL," & _
    "PRIMARY KEY (codigo))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTemp(" & _
    "codrelinp NUMERIC NULL," & _
    "idld NUMERIC NULL)"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbevotemp(" & _
    "campo01 VARCHAR(10) NOT NULL,campo02 VARCHAR(10) NOT NULL," & _
    "campo03 VARCHAR(50) NULL,campo04 VARCHAR(50) NULL," & _
    "campo05 VARCHAR(5) NULL,campo06 VARCHAR(50) NULL," & _
    "campo07 VARCHAR(15) NULL,campo08 VARCHAR(300) NULL," & _
    "campo09 VARCHAR(100) NOT NULL,campo10 VARCHAR(100) NULL," & _
    "campo11 VARCHAR(50) NULL," & _
    "fase01 VARCHAR(50) NULL,fase02 VARCHAR(50) NULL,fase03 VARCHAR(50) NULL," & _
    "fase04 VARCHAR(50) NULL,fase05 VARCHAR(50) NULL,fase06 VARCHAR(50) NULL," & _
    "fase07 VARCHAR(50) NULL,fase08 VARCHAR(50) NULL,fase09 VARCHAR(50) NULL," & _
    "fase10 VARCHAR(50) NULL,fase11 VARCHAR(50) NULL,fase12 VARCHAR(50) NULL," & _
    "fase13 VARCHAR(50) NULL,fase14 VARCHAR(50) NULL,fase15 VARCHAR(50) NULL," & _
    "fase16 VARCHAR(50) NULL,fase17 VARCHAR(50) NULL,fase18 VARCHAR(50) NULL," & _
    "fase19 VARCHAR(50) NULL,fase20 VARCHAR(50) NULL,fase21 VARCHAR(50) NULL," & _
    "fase22 VARCHAR(50) NULL,fase23 VARCHAR(50) NULL,fase24 VARCHAR(50) NULL," & _
    "fase25 VARCHAR(50) NULL,fase26 VARCHAR(50) NULL,fase27 VARCHAR(50) NULL," & _
    "fase28 VARCHAR(50) NULL,fase29 VARCHAR(50) NULL,fase30 VARCHAR(50) NULL," & _
    "fase31 VARCHAR(50) NULL,fase32 VARCHAR(50) NULL,fase33 VARCHAR(50) NULL," & _
    "fase34 VARCHAR(50) NULL,fase35 VARCHAR(50) NULL,fase36 VARCHAR(50) NULL," & _
    "fase37 VARCHAR(50) NULL,fase38 VARCHAR(50) NULL,fase39 VARCHAR(50) NULL," & _
    "fase40 VARCHAR(50) NULL,fase41 VARCHAR(50) NULL,fase42 VARCHAR(50) NULL," & _
    "fase43 VARCHAR(50) NULL,fase44 VARCHAR(50) NULL,fase45 VARCHAR(50) NULL," & _
    "fase46 VARCHAR(50) NULL,fase47 VARCHAR(50) NULL,fase48 VARCHAR(50) NULL," & _
    "fase49 VARCHAR(50) NULL,fase50 VARCHAR(50) NULL," & _
    "PRIMARY KEY (campo01, campo02, campo09))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTerceirizados(" & _
    "chapa VARCHAR(16) NOT NULL," & _
    "nome VARCHAR(100) NOT NULL," & _
    "foto TEXT NOT NULL," & _
    "idsetor VARCHAR(20) NOT NULL," & _
    "setor VARCHAR(100) NOT NULL," & _
    "idfuncao VARCHAR(20) NOT NULL," & _
    "funcao VARCHAR(100) NOT NULL," & _
    "idcc VARCHAR(50) NOT NULL," & _
    "nmcc VARCHAR(100) NOT NULL," & _
    "empresa VARCHAR(100) NOT NULL," & _
    "ativo CHAR(1) NOT NULL," & _
    "datacadastro DATETIME NOT NULL," & _
    "datacontratoini DATETIME NOT NULL," & _
    "datacontratofim DATETIME NULL," & _
    "PRIMARY KEY (chapa))"


'============================
    'TABELAS PADRAO
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbparametros(" & _
    "mediaaprovacao FLOAT NOT NULL," & _
    "geratreiint VARCHAR(1) NOT NULL," & _
    "aprovadorest FLOAT NOT NULL," & _
    "ativalog VARCHAR(1) NULL," & _
    "geratreiobr VARCHAR(1) NULL," & _
    "integrar VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "avisos VARCHAR(1) NULL," & _
    "atuautomatica VARCHAR(1) NULL," & _
    "caminhoexeauto VARCHAR(300) NULL," & _
    "calcexp VARCHAR(1) NULL," & _
    "afastdias VARCHAR(15) NULL," & _
    "afasttreiint VARCHAR(1) NULL," & _
    "afasttreiobr VARCHAR(1) NULL," & _
    "PRIMARY KEY (codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbintegracao(" & _
    "tipobanco NUMERIC NOT NULL," & _
    "sistema NUMERIC NOT NULL," & _
    "modulo CHAR(10) NOT NULL," & _
    "nserver VARCHAR(50) NULL," & _
    "nbanco VARCHAR(50) NULL," & _
    "nusuario VARCHAR(50) NULL," & _
    "nsenha VARCHAR(50) NULL," & _
    "codcoligada INT NULL," & _
    "PRIMARY KEY (tipobanco,sistema,modulo))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDadosEmpresa(" & _
    "razaosocial VARCHAR(100) NULL," & _
    "endereco VARCHAR(100) NULL," & _
    "bairro VARCHAR(50) NULL," & _
    "cidade VARCHAR(50) NULL," & _
    "uf VARCHAR(2) NULL," & _
    "cep VARCHAR(10) NULL," & _
    "email VARCHAR(100) NULL," & _
    "site VARCHAR(100) NULL," & _
    "telefone VARCHAR(20) NULL," & _
    "fax VARCHAR(20) NULL," & _
    "cnpj VARCHAR(30) NULL," & _
    "ie VARCHAR(30) NULL," & _
    "logo VARCHAR(300) NULL," & _
    "codcoligada INT NOT NULL," & _
    "ativa VARCHAR(1) NULL," & _
    "PRIMARY KEY (codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfEmail(" & _
    "smtp VARCHAR(100) NULL," & _
    "usuario VARCHAR(50) NULL," & _
    "senha VARCHAR(30) NULL," & _
    "codcoligada INT NULL)"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbLog(" & _
    "data VARCHAR(20) NULL," & _
    "hora VARCHAR(20) NOT NULL," & _
    "usuario VARCHAR(50) NOT NULL," & _
    "grupo VARCHAR(50) NULL," & _
    "formulario VARCHAR(50) NULL," & _
    "acao VARCHAR(300) NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfLV(" & _
    "nmusuario VARCHAR(50) NULL," & _
    "idmodulo NUMERIC NOT NULL," & _
    "indice NUMERIC NOT NULL," & _
    "posicao NUMERIC NOT NULL," & _
    "largura FLOAT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "PRIMARY KEY (id))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfGrupo(" & _
    "idgrupo NUMERIC NOT NULL," & _
    "idmenu NUMERIC NOT NULL," & _
    "idsub VARCHAR(10) NOT NULL," & _
    "tipo VARCHAR(20) NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "status VARCHAR(1) NOT NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "icon NUMERIC NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbSenha(" & _
    "usuario VARCHAR(50) NOT NULL," & _
    "senha VARCHAR(50) NOT NULL," & _
    "codigo NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMenu(" & _
    "idmenu NUMERIC NULL," & _
    "idsub VARCHAR(10) NULL," & _
    "tipo VARCHAR(20) NULL," & _
    "nome VARCHAR(50) NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
       
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbUsuarios(" & _
    "codigo NUMERIC NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "endereco VARCHAR(50) NULL," & _
    "cep VARCHAR(50) NULL," & _
    "bairro VARCHAR(50) NULL," & _
    "cidade VARCHAR(50) NULL," & _
    "uf VARCHAR(50) NULL," & _
    "fone VARCHAR(50) NULL," & _
    "celular VARCHAR(50) NULL," & _
    "ramal VARCHAR(50) NULL," & _
    "email VARCHAR(50) NULL," & _
    "codgrupo NUMERIC NULL," & _
    "altlogin NUMERIC NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "multiplic CHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbGrupo(" & _
    "codigo NUMERIC NOT NULL," & _
    "descricao VARCHAR(50) NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTipoFCE(" & _
    "id INT NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "descricao VARCHAR(300) NOT NULL," & _
    "ativo CHAR(1) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "cor VARCHAR(30) NOT NULL," & _
    "PRIMARY KEY (id))"
    
    
    'ABAIXO: CRIA CONFIGURAÇÃO PARA USUÁRIO ADMINISTRADOR
    oConn.Close
    
    vCodcoligada = 5 'Primeiro cadastro de coligada
    oConn.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=True;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName

    SqlSenha = "Insert into tbSenha(usuario,senha,codigo,codcoligada) Values('adm','123',1,'" & vCodcoligada & "');"
    rsSenha.Open SqlSenha, oConn
    
    SqlUsuario = "Insert into tbUsuarios(codigo,nome,codgrupo,ativo,codcoligada) Values(1,'Administrador do sistema',1,'S','" & vCodcoligada & "');"
    rsUsuario.Open SqlUsuario, oConn
    
    SqlGrupo = "Insert into tbGrupo(codigo,descricao,ativo,codcoligada) Values(1,'Administradores','S','" & vCodcoligada & "');"
    rsGrupo.Open SqlGrupo, oConn
    
    
    SqlConfGrupo = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'01','TAB','Cadastros','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'01','CAT','Primários','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'02','CAT','Secundários','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0101','BUT','Ramo de atividades','S','" & vCodcoligada & "',1);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0102','BUT','Clientes','S','" & vCodcoligada & "',2);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0103','BUT','Transportadoras','S','" & vCodcoligada & "',3);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0104','BUT','Tipo material','S','" & vCodcoligada & "',4);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0205','BUT','Materiais','S','" & vCodcoligada & "',5);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0206','BUT','Itens verificação','S','" & vCodcoligada & "',6);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0207','BUT','Projetos','S','" & vCodcoligada & "',7);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0208','BUT','Processos','S','" & vCodcoligada & "',8);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,2,'02','TAB','Orçamentos','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,2,'11','CAT','Vendas','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,2,'1111','BUT','Serviços','S','" & vCodcoligada & "',9);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'03','TAB','Planejamento','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'21','CAT','Planejamento e Controle de Produção','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'2121','BUT','FCE','S','" & vCodcoligada & "',10);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'2122','BUT','LM','S','" & vCodcoligada & "',11);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'2123','BUT','LD','S','" & vCodcoligada & "',12);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'2124','BUT','OS','S','" & vCodcoligada & "',13);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,3,'2125','BUT','Controle de Desenhos','S','" & vCodcoligada & "',28);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,4,'04','TAB','Produção','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,4,'31','CAT','Acompanhamento de Produção','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,4,'3131','BUT','OS Acompamenhamento','S','" & vCodcoligada & "',13);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,4,'3132','BUT','Evolução','S','" & vCodcoligada & "',14);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,5,'05','TAB','Inspeção/Expedição','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,5,'41','CAT','Emissão de Relatórios','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,5,'4141','BUT','Emitir Relatório','S','" & vCodcoligada & "',15);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,5,'4142','BUT','Imprimir relatório','S','" & vCodcoligada & "',16);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'06','TAB','Configurações','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'51','CAT','Parametrizações','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'52','CAT','Aparência','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5151','BUT','Sistema','S','" & vCodcoligada & "',17);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5152','BUT','Grupos','S','" & vCodcoligada & "',18);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5153','BUT','Usuários','S','" & vCodcoligada & "',19);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5254','BUT','Menu','S','" & vCodcoligada & "',20);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5255','BUT','Skin','S','" & vCodcoligada & "',21);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5256','BUT','Fundo','S','" & vCodcoligada & "',22);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'07','TAB','Sobre','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'61','CAT','Sobre','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'6161','BUT','Sobre ZEUS','S','" & vCodcoligada & "',23);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'6162','BUT','Ajuda do ZEUS','S','" & vCodcoligada & "',24);"
    
    rsConfGrupo.Open SqlConfGrupo, oConn
    
    SqlConfGrupo = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKINC','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKEDI','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKSAL','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKEXC','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKIMP','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKFIL','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKAVA','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKADI','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKDEM','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKADIRES','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,0,0,'CHK','CHKADIREP','S'," & vCodcoligada & ");"
    rsConfGrupo.Open SqlConfGrupo, oConn
    oConn.Close
    Set oConn = Nothing
       
    Msgbox "Tabelas criadas com sucesso", vbInformation, "ZEUS"
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Function

Public Function DesabBotoesN1(Frm As Form)
    Dim X As Integer
    For X = 0 To Frm.cmdconsulta.Count - 1
        Frm.cmdconsulta(X).UseGreyscale = True
    Next
End Function

Public Function HabBotoesN1(Frm As Form)
    Dim X As Integer
    For X = 0 To Frm.cmdconsulta.Count - 1
        Frm.cmdconsulta(X).UseGreyscale = False
    Next
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'ROTINAS/FUNÇÕES DO LISTVIEW GENERICO - DAKI PARA BAIXO
Public Function MontaFiltro()
    If Formulario = "Movimentações - OS" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Clientes" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Transportadoras" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Fórmulas de Produtos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Todos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        'frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Orçamentos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "FCE" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "CD - Controle de Desenhos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Tipo de materiais" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Cadastro de Desenhos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "LM" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Programação" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Ativos pendentes"
        frmFiltro.Combo1.List(2) = "Ativos agendados"
        frmFiltro.Combo1.List(3) = "Ativos concluidos"
        frmFiltro.Combo1.List(4) = "Ativos desmarcados"
        frmFiltro.Combo1.List(5) = "Cancelados"
        frmFiltro.Combo1.List(6) = "Todos"
        frmFiltro.Combo1.List(7) = "Programação"
        frmFiltro.Combo1.Text = "Ativos pendentes"
    End If
    If Formulario = "Fórmula - Centro de Custo" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "OS Fechamento - Permissão de Colaboradores" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Usuários" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Grupos" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "RNCF - Registro de Não Conformidade de Fabricação" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "CD - Comunicação de Desvio"
        frmFiltro.Combo1.List(1) = "CODAC - Coleta de Dados e Análise de Causas"
        frmFiltro.Combo1.List(2) = "DAAC - Definição da Ação e Análise Concluida"
        frmFiltro.Combo1.List(3) = "EVA - Execução e Verificação da Ação"
        frmFiltro.Combo1.List(4) = "TAC - Tomada de Ação Concluida"
        frmFiltro.Combo1.List(5) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Métodos & Processos" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Planejamento"
        'frmFiltro.Combo1.List(1) = "Produção"
        'frmFiltro.Combo1.List(2) = "Expedição"
        'frmFiltro.Combo1.List(0) = "Todos"
        'frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Relatórios de Inspeção" Then
        'TiPo = False
'        frmFiltro.Combo1.List(0) = "Avaliados"
'        frmFiltro.Combo1.List(1) = "Não Avaliados"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Relatórios de Expedição" Then
        'TiPo = False
'        frmFiltro.Combo1.List(0) = "Avaliados"
'        frmFiltro.Combo1.List(1) = "Não Avaliados"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    
    If Formulario = "Relatórios de Expedição emitidos" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Relatórios de Inspeção emitidos" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Faturamento por FCE" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.List(1) = "Em aberto"
        frmFiltro.Combo1.Text = "Todos"
    End If
End Function

Public Function MontaCabLV(Cab0 As String, Cab1 As String, Cab2 As String, Cab3 As String, Cab4 As String, Cab5 As String, Cab6 As String, Cab7 As String, Cab8 As String, Cab9 As String, Cab10 As String, Cab11 As String, Cab12 As String, Cab13 As String, Cab14 As String, Cab15 As String, Cab16 As String, Cab17 As String, Cab18 As String, Cab19 As String, Cab20 As String)
    NomeColLV(0) = Cab0
    NomeColLV(1) = Cab1
    NomeColLV(2) = Cab2
    NomeColLV(3) = Cab3
    NomeColLV(4) = Cab4
    NomeColLV(5) = Cab5
    NomeColLV(6) = Cab6
    NomeColLV(7) = Cab7
    NomeColLV(8) = Cab8
    NomeColLV(9) = Cab9
    NomeColLV(10) = Cab10
    NomeColLV(11) = Cab11
    NomeColLV(12) = Cab12
    NomeColLV(13) = Cab13
    NomeColLV(14) = Cab14
    NomeColLV(15) = Cab15
    NomeColLV(16) = Cab16
    NomeColLV(17) = Cab17
    NomeColLV(18) = Cab18
    NomeColLV(19) = Cab19
    NomeColLV(20) = Cab20
End Function

Public Function DimensionaLV(NomeLV As String)
    MeuLV.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
    MeuLV.Frame1.Caption = NomeLV
    MeuLV.Frame1.Move 0, 0, Principal.ScaleWidth - 300, Principal.ScaleHeight - 650
    MeuLV.ListView1.Move 100, 1000, Principal.ScaleWidth - 500, Principal.ScaleHeight - 1800
End Function

Public Function DimensionaForm()
    Dim XSize As Integer
    frmMonitorar.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
    frmMonitorar.Frame2.Move 5000, 120, Principal.ScaleWidth - 5580, Principal.ScaleHeight - 1000
    frmMonitorar.ListView1.Move 260, 500, Principal.ScaleWidth - 11650, Principal.ScaleHeight - 2200
    
    frmMonitorar.ListView4.Move 260, 240, Principal.ScaleWidth - 11650
    
    If FimAprop = "N" Then
        frmMonitorar.ListView3.Move 120 + Principal.ScaleWidth - 11300, 240, frmMonitorar.Frame12.Width + 700, Principal.ScaleHeight - 1200
    Else
        frmMonitorar.ListView3.Move 120 + Principal.ScaleWidth - 11300, 240, frmMonitorar.Frame12.Width + 700, Principal.ScaleHeight - 2000
    End If
    frmMonitorar.Command1.Move Principal.ScaleWidth - 3300, Principal.ScaleHeight - 1500
End Function

Public Function DimensionaPPS()
    frmProgramacao.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
    frmProgramacao.ListView1.Move 120, frmProgramacao.Top + 10, Principal.ScaleWidth - 500, Principal.ScaleHeight - 1800
    frmProgramacao.Frame1.Move 7020, frmProgramacao.Height - 1700, 3975, 1095
    frmProgramacao.Frame2.Move 11100, frmProgramacao.Height - 1700, 6615, 1095
    frmProgramacao.Frame3.Move 4400, frmProgramacao.Height - 1700, 2535, 1095
    frmProgramacao.cmdCadastro(12).Move 120, frmProgramacao.Height - 1350, 615, 615
    frmProgramacao.cmdCadastro(13).Move 735, frmProgramacao.Height - 1350, 615, 615
    frmProgramacao.cmdCadastro(0).Move 1350, frmProgramacao.Height - 1350, 615, 615
    frmProgramacao.Frame4.Move 2120, frmProgramacao.Height - 1700, 1935, 1095
    
    frmProgramacao.Frame5.Move 17800, frmProgramacao.Height - 1700, 3935, 1095
    
End Function

Public Function DimensionaFormInsp()
    Dim XSize As Integer
    chamaForm.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
    
    chamaForm.Frame3.Move chamaForm.Frame3.Left, chamaForm.Frame3.Top, chamaForm.Width - 6900, chamaForm.Height - 1655
    
    chamaForm.ListView1.Move chamaForm.ListView1.Left, chamaForm.ListView1.Top, chamaForm.Frame3.Width - 300, chamaForm.Frame3.Height - 800
    
    chamaForm.SkinLabel8.Move chamaForm.SkinLabel8.Left, chamaForm.Frame3.Height - 400, chamaForm.SkinLabel8.Width, chamaForm.SkinLabel8.Height
    chamaForm.SkinLabel29.Move chamaForm.SkinLabel29.Left, chamaForm.Frame3.Height - 400, chamaForm.SkinLabel29.Width, chamaForm.SkinLabel29.Height
    chamaForm.SkinLabel9.Move chamaForm.SkinLabel9.Left, chamaForm.Frame3.Height - 400, chamaForm.SkinLabel9.Width, chamaForm.SkinLabel9.Height
    chamaForm.SkinLabel30.Move chamaForm.SkinLabel30.Left, chamaForm.Frame3.Height - 400, chamaForm.SkinLabel30.Width, chamaForm.SkinLabel30.Height
    chamaForm.SkinLabel10.Move chamaForm.SkinLabel10.Left, chamaForm.Frame3.Height - 400, chamaForm.SkinLabel10.Width, chamaForm.SkinLabel10.Height
    
    
    chamaForm.Frame1.Move chamaForm.Frame1.Left, chamaForm.Frame1.Top, chamaForm.Frame1.Width, chamaForm.Height - 2485
    
    chamaForm.cmdCadastro(4).Move chamaForm.cmdCadastro(4).Left, chamaForm.Frame3.Height + 250
    chamaForm.cmdCadastro(6).Move chamaForm.cmdCadastro(6).Left, chamaForm.Frame3.Height + 250
End Function

Public Function DimensionaFormExp(vForm As Form)
    Dim XSize As Integer
    vForm.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
    
    vForm.Frame3.Move vForm.Frame3.Left, vForm.Frame3.Top, vForm.Width - 7700, vForm.Height - 1655
    'If vForm.txtcadastro(16).Top < vForm.Frame5.Height Then
        vForm.Frame4.Move vForm.Frame4.Left, vForm.Frame4.Top, vForm.Frame4.Width, vForm.Height - 4775
    'End If
    
    If vForm.Name = "frmRelExpAvulso" Then
        vForm.ListView1.Move vForm.ListView1.Left, vForm.ListView1.Top, vForm.Frame3.Width - 300, vForm.Frame3.Height - 2600
        vForm.Text1.Move vForm.Text1.Left, vForm.Text1.Top, vForm.Frame3.Width - 4600, vForm.Text1.Height
        
        vForm.cmdExpAvulso(8).Move vForm.Frame3.Width - 4400, vForm.cmdExpAvulso(8).Top, vForm.cmdExpAvulso(8).Width, vForm.cmdExpAvulso(8).Height
        vForm.Text2.Move vForm.Frame3.Width - 3900, vForm.Text2.Top, vForm.Text2.Width, vForm.Text2.Height
        vForm.SkinLabel27.Move vForm.Frame3.Width - 3900, vForm.SkinLabel27.Top, vForm.SkinLabel27.Width, vForm.SkinLabel27.Height
        vForm.SkinLabel28.Move vForm.Frame3.Width - 2500, vForm.SkinLabel28.Top, vForm.SkinLabel28.Width, vForm.SkinLabel28.Height
        vForm.Combo1.Move vForm.Frame3.Width - 2500, vForm.Combo1.Top, vForm.Combo1.Width
        vForm.txtCadastro(5).Move vForm.txtCadastro(5).Left, vForm.txtCadastro(5).Top, vForm.txtCadastro(5).Width, 330
    Else
        vForm.ListView1.Move vForm.ListView1.Left, vForm.ListView1.Top, vForm.Frame3.Width - 300, vForm.Frame3.Height - 1200
        vForm.SkinLabel9.Move vForm.SkinLabel9.Left, vForm.Frame3.Height - 800, vForm.SkinLabel9.Width, vForm.SkinLabel9.Height
        vForm.SkinLabel30.Move vForm.SkinLabel30.Left, vForm.Frame3.Height - 800, vForm.SkinLabel30.Width, vForm.SkinLabel30.Height
        vForm.SkinLabel10.Move vForm.SkinLabel10.Left, vForm.Frame3.Height - 800, vForm.SkinLabel10.Width, vForm.SkinLabel10.Height
        vForm.SkinLabel4.Move vForm.SkinLabel4.Left, vForm.Frame3.Height - 400, vForm.SkinLabel4.Width, vForm.SkinLabel4.Height
        vForm.txtCadastro(17).Move vForm.txtCadastro(17).Left, vForm.Frame3.Height - 450, vForm.txtCadastro(17).Width, vForm.txtCadastro(17).Height
    End If

    vForm.SkinLabel8.Move vForm.SkinLabel8.Left, vForm.Frame3.Height - 800, vForm.SkinLabel8.Width, vForm.SkinLabel8.Height
    vForm.SkinLabel29.Move vForm.SkinLabel29.Left, vForm.Frame3.Height - 800, vForm.SkinLabel29.Width, vForm.SkinLabel29.Height
    vForm.cmdCadastro(4).Move vForm.cmdCadastro(4).Left, vForm.Frame4.Height + 3350
    vForm.cmdCadastro(6).Move vForm.cmdCadastro(6).Left, vForm.Frame4.Height + 3350
'    If vForm.txtcadastro(16).Top > vForm.Frame5.Height Then
    If vForm.Frame4.Height < 4575 Then
        vForm.Frame4.Move vForm.Frame4.Left, vForm.Frame4.Top, vForm.Frame4.Width, vForm.Height - 4000
        vForm.cmdCadastro(4).Move vForm.Frame3.Left, vForm.Frame4.Height + 3350
        vForm.cmdCadastro(6).Move vForm.Frame3.Left + 615, vForm.Frame4.Height + 3350
        vForm.cmdCadastro(4).Move vForm.cmdCadastro(4).Left, vForm.Frame3.Height + 300
        vForm.cmdCadastro(6).Move vForm.cmdCadastro(6).Left, vForm.Frame3.Height + 300
    End If
End Function

Public Function MontaCabecalhoLV()
    Dim X As Integer
    'Limpa o cabeçalho antes de compor novamente
    MeuLV.ListView1.ColumnHeaders.Clear
    With MeuLV.ListView1.ColumnHeaders
        For X = 0 To 20
            If NomeColLV(X) = "" Then Exit Function
            .Add , , NomeColLV(X), Len(NomeColLV(X)) * 144
            QtdColReal = QtdColReal + 1
        Next
    End With
End Function

Public Function MontaCabecalhoLVTeste(vListView As Listview)
    Dim X As Integer
    'Limpa o cabeçalho antes de compor novamente
    vListView.ColumnHeaders.Clear
    With vListView.ColumnHeaders
        For X = 0 To 20
            If NomeColLV(X) = "" Then Exit Function
            .Add , , NomeColLV(X), Len(NomeColLV(X)) * 144
            QtdColReal = QtdColReal + 1
        Next
    End With
End Function


Public Function MontaDadosLV(ZeroPriCol As String)
On Error GoTo Err
    ' Declaração de variaveis
    Dim rsListview As New ADODB.Recordset ' Variavel que vai receber os dados da tabela
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim X As Integer, Y As Integer
    rsListview.Open SqlLV, cnBanco, adOpenKeyset, adLockReadOnly
    MeuLV.ListView1.ListItems.Clear 'Limpa o listview
    If rsListview.RecordCount <> 0 Then Principal.ProgressBar1.Max = rsListview.RecordCount
    X = 0
    While Not rsListview.EOF
        Y = 0
        Principal.ProgressBar1.Value = X
        If ZeroPriCol = "S" Then
            If Not IsNull(rsListview(Y)) Then Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(rsListview(Y), "000000")) Else Set ItemLst = MeuLV.ListView1.ListItems.Add(, , "-")
        Else
            If Not IsNull(rsListview(Y)) Then Set ItemLst = MeuLV.ListView1.ListItems.Add(, , rsListview(Y)) Else Set ItemLst = MeuLV.ListView1.ListItems.Add(, , "-")
        End If
        For Y = 1 To QtdColReal - 1
            If Not IsNull(rsListview.Fields(Y)) Then ItemLst.SubItems(Y) = rsListview.Fields(Y) Else ItemLst.SubItems(Y) = "-"
        Next
        rsListview.MoveNext
        X = X + 1
    Wend
    'NAO EXECUTAR A LINHA ABAIXO AO ENTRAR NO FILTRO
    If vControlaDim < 8 Then LV_AutoSizeColumn MeuLV.ListView1
    vControlaDim = vControlaDim + 1
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    rsListview.Close
    Set rsListview = Nothing
    Principal.StatusBar1.Panels(5).Text = "Registros: " & MeuLV.ListView1.ListItems.Count
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Exit Function
    End If
End Function

Public Function MontaDadosLVTeste(ZeroPriCol As String, vListView As Listview)
On Error GoTo Err
    ' Declaração de variaveis
    Dim rsListview As New ADODB.Recordset ' Variavel que vai receber os dados da tabela
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim X As Integer, Y As Integer
    rsListview.Open SqlLV, cnBanco, adOpenKeyset, adLockReadOnly
    vListView.ListItems.Clear 'Limpa o listview
    If rsListview.RecordCount <> 0 Then Principal.ProgressBar1.Max = rsListview.RecordCount
    X = 0
    While Not rsListview.EOF
        Y = 0
        Principal.ProgressBar1.Value = X
        If ZeroPriCol = "S" Then
            If Not IsNull(rsListview(Y)) Then Set ItemLst = vListView.ListItems.Add(, , Format(rsListview(Y), "000000")) Else Set ItemLst = vListView.ListItems.Add(, , "-")
        Else
            If Not IsNull(rsListview(Y)) Then Set ItemLst = vListView.ListItems.Add(, , rsListview(Y)) Else Set ItemLst = vListView.ListItems.Add(, , "-")
        End If
        For Y = 1 To QtdColReal - 1
            If Not IsNull(rsListview.Fields(Y)) Then ItemLst.SubItems(Y) = rsListview.Fields(Y) Else ItemLst.SubItems(Y) = "-"
        Next
        rsListview.MoveNext
        X = X + 1
    Wend
    'NAO EXECUTAR A LINHA ABAIXO AO ENTRAR NO FILTRO
    If vControlaDim < 8 Then LV_AutoSizeColumnTeste vListView
    vControlaDim = vControlaDim + 1
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    rsListview.Close
    Set rsListview = Nothing
    Principal.StatusBar1.Panels(5).Text = "Registros: " & vListView.ListItems.Count
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Exit Function
    End If
End Function

Public Function carregaCoresTipoFCE()
On Error GoTo Err
    Dim rsCorTipoFCE As New ADODB.Recordset
    Dim sqlCorTipoFCE As String
    Dim X As Integer
    sqlCorTipoFCE = "select nome,cor from tbTipoFCE"
    rsCorTipoFCE.Open sqlCorTipoFCE, cnBanco, adOpenKeyset, adLockReadOnly
    For X = 0 To rsCorTipoFCE.RecordCount - 1
        vCorTipoFCE(X, 0) = rsCorTipoFCE.Fields(0)
        vCorTipoFCE(X, 1) = rsCorTipoFCE.Fields(1)
        rsCorTipoFCE.MoveNext
    Next
    rsCorTipoFCE.Close
    Set rsCorTipoFCE = Nothing
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

Public Function MudaPropPicture()
    MeuLV.picBg.BackColor = MeuLV.ListView1.BackColor
    MeuLV.picBg.ScaleMode = vbTwips
    MeuLV.picBg.BorderStyle = vbBSNone
    MeuLV.picBg.AutoRedraw = True
    MeuLV.picBg.Visible = False
End Function

'MUDA COR DAS LINHAS DO LISTVIEW PARA IDENTIFICAR O TIPO DE FCE
'Public Function MudaCorList(Listview As Listview, coluna As Integer, colorirLinha As Integer)
Public Function MudaCorList(Listview As Listview)
    Dim i As Integer
    MeuLV.picBg.Width = Listview.Width
    MeuLV.picBg.Height = Listview.ListItems(1).Height * (Listview.ListItems.Count)
    MeuLV.picBg.ScaleHeight = Listview.ListItems.Count
    MeuLV.picBg.ScaleWidth = 1
    MeuLV.picBg.DrawWidth = 1
    MeuLV.picBg.Cls
    Listview.BackColor = &H80000018
    
    For i = 1 To 5000
        'If vGuardaLinhaTipo(i, 0) = "" Then Exit For
        localizaCorTipoFCE Trim(vGuardaLinhaTipo(i, 0))
        If vIDCorTipoFCE <> "" Then MeuLV.picBg.Line (0, vGuardaLinhaTipo(i, 1) - 1)-(1, vGuardaLinhaTipo(i, 1)), vIDCorTipoFCE, BF
    Next
    
    'If Trim(Listview.ListItems.Item(colorirLinha).SubItems(coluna)) <> "-" Then
    '    localizaCorTipoFCE Trim(Listview.ListItems.Item(colorirLinha).SubItems(coluna))
    '    If vIDCorTipoFCE <> "" Then MeuLV.picBg.Line (0, colorirLinha - 1)-(1, colorirLinha), vIDCorTipoFCE, BF
    'End If
    Listview.Picture = MeuLV.picBg.Image
    Listview.Refresh
End Function

Private Function guardaLinhaTipo(Listview As Listview, vColuna As Integer, vLinha As Integer)
    'PRIMEIRO DEVE LIMPAR OS DADOS
    
    vGuardaLinhaTipo(vLinha, 0) = Trim(Listview.ListItems.Item(vLinha).SubItems(vColuna))
    vGuardaLinhaTipo(vLinha, 1) = vLinha
End Function

Public Function limpaGuardaLinhaTipo()
    Dim X As Integer, Y As Integer
'    For X = LBound(vQualquerDado) To UBound(vQualquerDado)
    For X = 0 To 5000
        For Y = 0 To 1
            vGuardaLinhaTipo(X, Y) = ""
        Next
    Next X
End Function


Public Function PersonaColLV(posCol As Integer, negritoCol As String, corCol As String, caracterCol As String, imageCol As String, formataColZero As String, formataColDecimal As String, alinhaCol As String)
    'On Error Resume Next
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        Principal.ProgressBar1.Value = X
        Set ItemLst = MeuLV.ListView1.ListItems.Item(X)
        'NEGRITO NOS ITENS DA COLUNA
        If negritoCol = "S" Then ItemLst.ListSubItems(posCol).Bold = True
        
        If corCol = "P" And ItemLst.ListSubItems(posCol) <> "-" Then
            'MUDA COR DAS LINHAS DO LISTVIEW PARA IDENTIFICAR O TIPO DE FCE
            'MudaCorList MeuLV.ListView1, posCol, X
            guardaLinhaTipo MeuLV.ListView1, posCol, X
        End If
        
        'COR VERDE/VERMELHO NOS ITENS DA COLUNA
        If corCol = "S" Then
            If Formulario = "ADP" Then
                If Date > CDate(ItemLst.ListSubItems(5)) And ItemLst.ListSubItems(9) <> "Concluido" Then
                    ItemLst.ForeColor = &HFF&
                    ItemLst.ListSubItems(1).ForeColor = &HFF&
                    ItemLst.ListSubItems(2).ForeColor = &HFF&
                    ItemLst.ListSubItems(3).ForeColor = &HFF&
                    ItemLst.ListSubItems(4).ForeColor = &HFF&
                    ItemLst.ListSubItems(5).ForeColor = &HFF&
                    ItemLst.ListSubItems(6).ForeColor = &HFF&
                    'ItemLst.ListSubItems(7).ForeColor = &HFF&
                    ItemLst.ListSubItems(8).ForeColor = &HFF&
                    ItemLst.ListSubItems(9).ForeColor = &HFF&
                    ItemLst.ListSubItems(10).ForeColor = &HFF&
                    ItemLst.ListSubItems(11).ForeColor = &HFF&
                    ItemLst.ListSubItems(12).ForeColor = &HFF&
                End If
                If ItemLst.ListSubItems(7) >= IniciaRelsEm Then
                    ItemLst.ListSubItems(7).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(7) < IniciaRelsEm And ItemLst.ListSubItems(7) >= vAprovadoRest Then
                    ItemLst.ListSubItems(7).ForeColor = &H80FF&
                Else
                    ItemLst.ListSubItems(7).ForeColor = &HC0&
                End If
            ElseIf Formulario = "Métodos & Processos" Then
                If ItemLst.ListSubItems(posCol) = "Planejamento" Then 'PRETO
                    ItemLst.ForeColor = &H80000008
                    ItemLst.ListSubItems(1).ForeColor = &H80000008
                    ItemLst.ListSubItems(2).ForeColor = &H80000008
                    ItemLst.ListSubItems(3).ForeColor = &H80000008
                    ItemLst.ListSubItems(4).ForeColor = &H80000008
                    ItemLst.ListSubItems(5).ForeColor = &H80000008
                    ItemLst.ListSubItems(6).ForeColor = &H80000008
                    ItemLst.ListSubItems(7).ForeColor = &H80000008
                    ItemLst.ListSubItems(8).ForeColor = &H80000008
                    ItemLst.ListSubItems(9).ForeColor = &H80000008
                    ItemLst.ListSubItems(10).ForeColor = &H80000008
                    ItemLst.ListSubItems(11).ForeColor = &H80000008
                    ItemLst.ListSubItems(12).ForeColor = &H80000008
                ElseIf ItemLst.ListSubItems(posCol) = "Produção" Then 'VERDE
                    ItemLst.ForeColor = &H8000&
                    ItemLst.ListSubItems(1).ForeColor = &H8000&
                    ItemLst.ListSubItems(2).ForeColor = &H8000&
                    ItemLst.ListSubItems(3).ForeColor = &H8000&
                    ItemLst.ListSubItems(4).ForeColor = &H8000&
                    ItemLst.ListSubItems(5).ForeColor = &H8000&
                    ItemLst.ListSubItems(6).ForeColor = &H8000&
                    ItemLst.ListSubItems(7).ForeColor = &H8000&
                    ItemLst.ListSubItems(8).ForeColor = &H8000&
                    ItemLst.ListSubItems(9).ForeColor = &H8000&
                    ItemLst.ListSubItems(10).ForeColor = &H8000&
                    ItemLst.ListSubItems(11).ForeColor = &H8000&
                    ItemLst.ListSubItems(12).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(posCol) = "Expedição" Then 'CINZA
                    ItemLst.ForeColor = &H808080
                    ItemLst.ListSubItems(1).ForeColor = &H808080
                    ItemLst.ListSubItems(2).ForeColor = &H808080
                    ItemLst.ListSubItems(3).ForeColor = &H808080
                    ItemLst.ListSubItems(4).ForeColor = &H808080
                    ItemLst.ListSubItems(5).ForeColor = &H808080
                    ItemLst.ListSubItems(6).ForeColor = &H808080
                    ItemLst.ListSubItems(7).ForeColor = &H808080
                    ItemLst.ListSubItems(8).ForeColor = &H808080
                    ItemLst.ListSubItems(9).ForeColor = &H808080
                    ItemLst.ListSubItems(10).ForeColor = &H808080
                    ItemLst.ListSubItems(11).ForeColor = &H808080
                    ItemLst.ListSubItems(12).ForeColor = &H808080
                End If
            Else
                'If ItemLst.ListSubItems(posCol) >= IniciaRelsEm Then
                    ItemLst.ListSubItems(posCol).ForeColor = &H8000&
                'ElseIf ItemLst.ListSubItems(posCol) < IniciaRelsEm And ItemLst.ListSubItems(posCol) >= vAprovadoRest Then
                '    ItemLst.ListSubItems(posCol).ForeColor = &H80FF&
                'Else
                '    ItemLst.ListSubItems(posCol).ForeColor = &HC0&
                'End If
            End If
        End If
        'CASAS DECIMAIS NOS ITENS DA COLUNA
        If formataColDecimal = "S" Then ItemLst.SubItems(posCol) = "" & Format(ItemLst.SubItems(posCol), "#,##0.00;(#,##0.00)")
        'FORMATAÇÃO DE 6 ZEROS NOS ITENS DA COLUNA
        If formataColZero = "S" Then ItemLst.SubItems(posCol) = "" & Format(ItemLst.SubItems(posCol), "000000")
        'ADICIONAR CARACTER(ES) NOS ITENS DA COLUNA
        If caracterCol <> "" Then
            If ItemLst.ListSubItems(posCol) <> "-" Then
                ItemLst.SubItems(posCol) = caracterCol & ItemLst.ListSubItems(posCol)
            End If
        End If
        'INFORMA SE IRÁ UTILIZAR O IMAGELIST NOS ITENS DA COLUNA
        If imageCol = "S" Then
            'A condição abaixo verifica o conteudo da posição do Listview
            If ItemLst.SubItems(posCol) = "S" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
                If apontaLV = 3 Then ItemLst.ListSubItems.Item(posCol).ReportIcon = "EXC"
            ElseIf ItemLst.SubItems(posCol) = "1" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "NEGATIVO"
            ElseIf ItemLst.SubItems(posCol) = "2" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "POSITIVO"
            ElseIf ItemLst.SubItems(posCol) = "3" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "ARQUIVADO"
            ElseIf ItemLst.SubItems(posCol) = "4" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "AGUARDE-02"
            ElseIf ItemLst.SubItems(posCol) = "5" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
            ElseIf ItemLst.SubItems(posCol) = "6" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "PENDENTE1"
            ElseIf ItemLst.SubItems(posCol) = "7" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "AVALIANDO1"
            ElseIf ItemLst.SubItems(posCol) = "8" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "CONCLUIDO1"
            ElseIf ItemLst.SubItems(posCol) = "9" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "FABRICANDO"
            ElseIf ItemLst.SubItems(posCol) = "20" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "FECHADO"
            ElseIf ItemLst.SubItems(posCol) = "10" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "PRETO"
            ElseIf ItemLst.SubItems(posCol) = "ANDAMENTO" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "ANDAMENTO"
            ElseIf ItemLst.SubItems(posCol) = "CONCLUIDA" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "CONCLUIDA"
            ElseIf ItemLst.SubItems(posCol) = "PARALIZADA" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "PARALIZADA"
            ElseIf ItemLst.SubItems(posCol) = "DUVIDA" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "DUVIDA"
            Else
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "EXC"
                If apontaLV = 3 Then ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
            End If
            ItemLst.SubItems(posCol) = ""
        End If
        'ALINHAMENTO DA COLUNA
        If alinhaCol = "D" Then
            MeuLV.ListView1.ColumnHeaders(posCol + 1).Alignment = lvwColumnRight
        ElseIf alinhaCol = "E" Then
            MeuLV.ListView1.ColumnHeaders(posCol + 1).Alignment = lvwColumnLeft
        Else
            MeuLV.ListView1.ColumnHeaders(posCol + 1).Alignment = lvwColumnCenter
        End If
    Next
    
    If corCol = "P" Then
        MudaCorList MeuLV.ListView1
    End If
    limpaGuardaLinhaTipo
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Function


Public Sub ExcluirDadosLV()
On Error GoTo Err
    Dim ItemLst As ListItem
    Dim rsExcLVGeral As New ADODB.Recordset
10  cnBanco.BeginTrans
    mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
    If Tp = 1 Then
        'SqlExcLVGeral = "Delete from tbHabilidades where codHabilidade= " & Val(varGlobal)
        'rsExcLVGeral.Open SqlExcLVGeral, cnBanco
        mobjMsg.Abrir "Registro excluido com sucesso", Ok, informacao, "ZEUS"
        'rsExcLVGeral.Update
    End If
    cnBanco.CommitTrans
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Public Sub ExcluirItemLV(LV As Listview)
On Error Resume Next
    Dim X As Integer, Y As Integer
    Y = LV.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    LV.ListItems.Remove (X)
End Sub

'Ajusta automaticamente a largura das colunas
Public Sub LV_AutoSizeColumn(LV As Listview, Optional Column As ColumnHeader = Nothing)
On Error GoTo Err
    Dim c As ColumnHeader
    Dim posi As Integer
    posi = 1
    
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            MeuLV.ListView1.ListItems.Item(1).Selected = True
            If posi = 1 Then
                If Len(MeuLV.ListView1.ListItems.Item(posi)) < Len(c) Then
                    'SendMessage LV.hwnd, LVM_FIRST + 31, c.Index - 1, -1
                    
                    'If Mid$(MeuLV.ListView1.ColumnHeaders.Item(posi).Width, 1, 3) = 180 Then
                        MeuLV.ListView1.ColumnHeaders.Item(posi).Width = 1400
                    'End If
                End If
                posi = posi + 1
            Else
                If Len(MeuLV.ListView1.SelectedItem.ListSubItems.Item(posi - 1)) > Len(c) Then
                    SendMessage LV.HWnd, LVM_FIRST + 30, c.Index - 1, -1
                    
                    If Mid$(MeuLV.ListView1.ColumnHeaders.Item(posi).Width, 1, 3) = 180 Then
                        MeuLV.ListView1.ColumnHeaders.Item(posi).Width = 0
                    End If
                End If
                posi = posi + 1
            End If
        Next
    Else
        SendMessage LV.HWnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
    Exit Sub
Err:
    Resume Next
End Sub

Public Sub LV_AutoSizeColumnTeste(LV As Listview, Optional Column As ColumnHeader = Nothing)
On Error GoTo Err
    Dim c As ColumnHeader
    Dim posi As Integer
    posi = 1
    
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
           LV.ListItems.Item(1).Selected = True
            If posi = 1 Then
                If Len(LV.ListItems.Item(posi)) < Len(c) Then
                    'SendMessage LV.hwnd, LVM_FIRST + 31, c.Index - 1, -1
                    
                    'If Mid$(MeuLV.ListView1.ColumnHeaders.Item(posi).Width, 1, 3) = 180 Then
                        LV.ColumnHeaders.Item(posi).Width = 1400
                    'End If
                End If
                posi = posi + 1
            Else
                If Len(LV.SelectedItem.ListSubItems.Item(posi - 1)) > Len(c) Then
                    SendMessage LV.HWnd, LVM_FIRST + 30, c.Index - 1, -1
                    
                    If Mid$(LV.ColumnHeaders.Item(posi).Width, 1, 3) = 180 Then
                        LV.ColumnHeaders.Item(posi).Width = 0
                    End If
                End If
                posi = posi + 1
            End If
        Next
    Else
        SendMessage LV.HWnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
    Exit Sub
Err:
    Resume Next
End Sub


'ROTINAS/FUNÇÕES DO LISTVIEW GENERICO - DAKI PARA CIMA
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Avaliador(Tipo As String)
On Error GoTo Err
'On Error Resume Next
    Dim rsAvaliador As New ADODB.Recordset
    Dim SqlAvaliador As String
    Dim X As Integer
    Dim Contador As Double, ConverTido As Double
    
    chamaForm.mskCadMatriz.PromptInclude = False
    
    Dim PontosColabExp As Double
    Dim PontosTotaisHab As Double
    Dim PontosTotaisTrei As Double
    Dim PontosTotaisFor As Double
    Contador = 0
    
    If chamaForm.Caption = "Cadastro de colaboradores" Then
        For X = 0 To 4
            If chamaForm.chkAvaliador(X).Value = 1 Then
                Contador = Contador + 1
            End If
        Next
    Else
        For X = 0 To 3
            If chamaForm.chkAvaliador(X).Value = 1 Then
                Contador = Contador + 1
            End If
        Next
    End If
    chamaForm.Label37.Caption = ""
    chamaForm.Label38.Caption = ""
    chamaForm.Label39.Caption = ""
    chamaForm.Label40.Caption = ""
    chamaForm.Label41.Caption = ""
    If chamaForm.Caption = "Cadastro de colaboradores" Then
        chamaForm.Label43 = ""
    End If
    If Contador = 0 Then Exit Function
    'PRIMEIRO CHECKBOX - EXPERIENCIA
    If chamaForm.chkAvaliador(0).Value = 1 Then
        Dim PontosMatrizExp As Double
        Dim ContCargoMatExp As Double
        
        SqlAvaliador = "select * from tbMatrizExp as a left join tbColaboradoresExp as b on a.codcargo = b.codcargo and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and b.cpf = '" & chamaForm.mskCadMatriz & "' and     b.tipo = '" & Tipo & "' where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        ContCargoMatExp = 0
        PontosMatrizExp = 0
        PontosColabExp = 0
        '>>Soma todos os pontos de EXPERIENCIA da matriz
        If rsAvaliador.RecordCount > 0 Then
            If Mid$(rsAvaliador.Fields(2), 4, 4) = "Anos" Then
                'Converte anos para meses
                ConverTido = Val(Mid$(rsAvaliador.Fields(2), 1, 3)) * 12
            Else
                ConverTido = Val(Mid$(rsAvaliador.Fields(2), 1, 3))
            End If
        End If
        PontosMatrizExp = ConverTido 'valor em meses = 100%
        ContCargoMatExp = ContCargoMatExp + 1
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        ConverTido = 0
        While Not rsAvaliador.EOF
            '>>Soma todos os pontos de EXPERIENCIA do colaborador
            If Not IsNull(rsAvaliador.Fields(8)) Then
                If Mid$(rsAvaliador.Fields(8), 5, 4) = "Anos" Then
                    'Converte anos para meses
                    ConverTido = Val(Mid$(rsAvaliador.Fields(8), 1, 3)) * 12
                Else
                    ConverTido = Val(Mid$(rsAvaliador.Fields(8), 1, 3))
                End If
            Else
                ConverTido = 0
            End If
            
            If ConverTido > PontosMatrizExp Then ConverTido = PontosMatrizExp
            
            If PontosMatrizExp <> 0 Then
                PontosColabExp = PontosColabExp + ((ConverTido * 100) / PontosMatrizExp)
            Else
                PontosColabExp = 0
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            rsAvaliador.MoveNext
        Wend
        'para encontrar a pontuação em Experiencia do colaborador:
        'Divide-se a quantidade de pontos encontrados para o colaborador (PontosColabExp)
        'pela quantidade de cargos da matriz (ContCargoMatExp)
        PontosColabExp = PontosColabExp / ContCargoMatExp
        
        If PontosMatrizExp <= 0 Then PontosColabExp = 100
        If PontosColabExp >= IniciaRelsEm Then
            chamaForm.Label37.ForeColor = &H8000&
        ElseIf PontosColabExp < IniciaRelsEm And PontosColabExp >= vAprovadoRest Then
            chamaForm.Label37.ForeColor = &H80FF&
        Else
            chamaForm.Label37.ForeColor = &HC0&
        End If
        
        If PontosColabExp > 100 Then PontosColabExp = 100
        
        chamaForm.Label37 = Format(PontosColabExp, "#,##0.00;(#,##0.00)") & " %"
        rsAvaliador.Close
        Set rsAvaliador = Nothing
    End If
    'SEGUNDO CHECKBOX - HABILIDADES
    If chamaForm.chkAvaliador(1).Value = 1 Then
        Dim PontosMatrizHab As Double
        Dim PontosColabHab As Double
        SqlAvaliador = "select a.cpf,a.codmatriz,a.codhabilidade,a.pontuacao,b.peso from tbColaboradoresHab as a inner join tbHabilidades as b on a.codcoligada = '" & vCodcoligada & "' and a.codhabilidade = b.codhabilidade and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and a.cpf = '" & chamaForm.mskCadMatriz & "' and     a.tipo = '" & Tipo & "' where a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        PontosMatrizHab = 0
        PontosColabHab = 0
        While Not rsAvaliador.EOF
            PontosMatrizHab = PontosMatrizHab + rsAvaliador.Fields(4)
            PontosColabHab = PontosColabHab + rsAvaliador.Fields(3)
            rsAvaliador.MoveNext
        Wend
        If PontosColabHab > 0 Then PontosTotaisHab = (PontosColabHab * 100) / PontosMatrizHab
        If PontosTotaisHab >= IniciaRelsEm Then
            chamaForm.Label38.ForeColor = &H8000&
        ElseIf PontosTotaisHab < IniciaRelsEm And PontosTotaisHab >= vAprovadoRest Then
            chamaForm.Label38.ForeColor = &H80FF&
        Else
            chamaForm.Label38.ForeColor = &HC0&
        End If
        
        If PontosMatrizHab <= 0 Then PontosTotaisHab = 0
        
        chamaForm.Label38 = Format(PontosTotaisHab, "#,##0.00;(#,##0.00)") & " %"
        rsAvaliador.Close
        Set rsAvaliador = Nothing
    End If
    'TERCEIRO CHECKBOX - CURSOS/TREINAMENTOS
    If chamaForm.chkAvaliador(2).Value = 1 Then
        Dim PontosMatrizTrei As Double
        Dim PontosColabTrei As Double
        
'        SqlAvaliador = "select * from tbMatrizCur as a left join tbcolaboradoresCur as b on a.codtreinamento = b.codtreinamento and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & TiPo & "' where a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        SqlAvaliador = "select * from tbMatrizCur as a left join tbcolaboradoresCur as b on a.codtreinamento = b.codtreinamento and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & Tipo & "' and b.codnivel >= a.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        PontosMatrizTrei = 0
        PontosTotaisTrei = 0
        While Not rsAvaliador.EOF
            PontosMatrizTrei = PontosMatrizTrei + 1
            If Not IsNull(rsAvaliador.Fields(3)) And rsAvaliador.Fields(6) <> "SR" Then PontosColabTrei = PontosColabTrei + 1
            rsAvaliador.MoveNext
        Wend
        If PontosMatrizTrei > 0 Then
            PontosTotaisTrei = (PontosColabTrei * 100) / PontosMatrizTrei
            If PontosTotaisTrei >= IniciaRelsEm Then
                chamaForm.Label39.ForeColor = &H8000&
            ElseIf PontosTotaisTrei < IniciaRelsEm And PontosTotaisTrei >= vAprovadoRest Then
                chamaForm.Label39.ForeColor = &H80FF&
            Else
                chamaForm.Label39.ForeColor = &HC0&
            End If
            
            chamaForm.Label39 = Format(PontosTotaisTrei, "#,##0.00;(#,##0.00)") & " %"
        Else
            chamaForm.Label39.ForeColor = &H8000&
            If PontosMatrizTrei <= 0 Then PontosTotaisTrei = 100
            chamaForm.Label39 = Format(PontosTotaisTrei, "#,##0.00;(#,##0.00)") & " %"
        End If
        
        rsAvaliador.Close
        Set rsAvaliador = Nothing
    End If
    'QUARTO CHECKBOX - FORMAÇÃO ESCOLAR
    If chamaForm.chkAvaliador(3).Value = 1 Then
        Dim PontosColabFor As Double
        Dim VerificaNull As Integer
        SqlAvaliador = "select c.codmatriz,c.codescolaridade,c.pontuacao,b.cpf,b.tipo,b.codescolaridade,a.peso from tbescolaridade as a left join tbcolaboradoresesc as b on a.codescolaridade = b.codescolaridade and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & Tipo & "' left join tbmatrizEsc as c on a.codescolaridade = c.codescolaridade and c.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' where a.codcoligada = '" & vCodcoligada & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        PontosColabFor = 0
        VerificaNull = 0
        Do While Not rsAvaliador.EOF
            If Not IsNull(rsAvaliador.Fields(5)) Then VerificaNull = VerificaNull + 1
            If Not IsNull(rsAvaliador.Fields(2)) Then PontosColabFor = rsAvaliador.Fields(2)
            If Not IsNull(rsAvaliador.Fields(3)) Then
                Exit Do
            End If
            rsAvaliador.MoveNext
        Loop
        If VerificaNull = 0 Then PontosColabFor = 0
        If PontosColabFor >= IniciaRelsEm Then
            chamaForm.Label40.ForeColor = &H8000&
        ElseIf PontosColabFor < IniciaRelsEm And PontosColabFor >= vAprovadoRest Then
            chamaForm.Label40.ForeColor = &H80FF&
        Else
            chamaForm.Label40.ForeColor = &HC0&
        End If
        
        chamaForm.Label40 = Format(PontosColabFor, "#,##0.00;(#,##0.00)") & " %"
        
        rsAvaliador.Close
        Set rsAvaliador = Nothing
    
    End If
    
    
    'QUINTO CHECKBOX - AVALIAÇÃO DE DESEMPENHO PROFISSIONAL
    Dim PontosColabADP As Double
    If chamaForm.Caption = "Cadastro de colaboradores" Then
        If chamaForm.chkAvaliador(4).Value = 1 Then
            SqlAvaliador = "select a.codcolaborador, b.nomecolaborador, MAX(a.nota) as Nota, MAX(a.datadevolucao) as datadevolucao from tbListaADP as a inner join tbcolaboradores  as b on a.codcolaborador = b.id " & _
            "where b.codcolaborador = '" & chamaForm.txtCadMatriz(2) & "' group by a.codcolaborador,b.nomecolaborador"
            rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
            PontosColabADP = 0
            If Not rsAvaliador.EOF Then
                If IsNull(rsAvaliador.Fields(2)) Then
                    PontosColabADP = 0
                Else
                    PontosColabADP = rsAvaliador.Fields(2)
                End If
            End If
        
            If PontosColabADP >= IniciaRelsEm Then
                chamaForm.Label43.ForeColor = &H8000&
            ElseIf PontosColabADP < IniciaRelsEm And PontosColabADP >= vAprovadoRest Then
                chamaForm.Label43.ForeColor = &H80FF&
            Else
                chamaForm.Label43.ForeColor = &HC0&
            End If
        
            chamaForm.Label43 = Format(PontosColabADP, "#,##0.00;(#,##0.00)") & " %"
        
            rsAvaliador.Close
            Set rsAvaliador = Nothing
        End If
    End If
    
    'mskCadMatriz.PromptInclude = True
    If Contador > 0 Then
        If ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor + PontosColabADP) / Contador) >= IniciaRelsEm Then
            chamaForm.Label41.ForeColor = &H8000&
        ElseIf ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor) + PontosColabADP / Contador) < IniciaRelsEm And ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor) + PontosColabADP / Contador) >= vAprovadoRest Then
            chamaForm.Label41.ForeColor = &H80FF&
        Else
            chamaForm.Label41.ForeColor = &HC0&
        End If
        
        If Contador > 0 Then
            chamaForm.Label41 = Format(((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor + PontosColabADP) / Contador), "#,##0.00;(#,##0.00)") & " %"
        End If
    Else
        chamaForm.Label41.ForeColor = &HC0&
        chamaForm.Label41 = "00,00 %"
    End If
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Exit Function
    End If
End Function

Public Function gravaLog(Campo1 As String, campo2 As String, campo3 As String)
On Error GoTo Err
    If GeraLog = "N" Then Exit Function
    Dim sqlLog As String
    Dim rsLog As New ADODB.Recordset
    
    sqlLog = "Insert into tbLog(data,hora,usuario,grupo,formulario,acao,codcoligada) Values('" & CStr(Date) & "','" & CStr(Time) & "','" & NomUsu & "','" & GrupoUsu & "','" & Formulario & "','" & Pesquisa & ":" & Campo1 & "-" & campo2 & "-" & campo3 & "','" & vCodcoligada & "')"
    rsLog.Open sqlLog, cnBanco
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

Public Function gravaSolicitacao(vCPF As String, vTipo As String, vNota As String, vSolicitacao As String, vSolicitante As String)
On Error GoTo Err
    Dim sqlSolicita As String
    Dim rsSolicita As New ADODB.Recordset
    
    Dim sqlCandColab As String
    Dim rsCandColab As New ADODB.Recordset
    
    'Insere a solicitação na tabela tbAutorizacao
    sqlSolicita = "Insert into tbAutorizacao(cpf,tipo,nota,solicitacao,datasolicitacao,solicitante,codcoligada) Values('" & vCPF & "','" & vTipo & "','" & vNota & "','" & vSolicitacao & "', substring(Convert(Char, getdate(), 103), 1, 10),'" & vSolicitante & "','" & vCodcoligada & "')"
    rsSolicita.Open sqlSolicita, cnBanco
    
    'retorna o último valor de identidade gerado para a tabela tbAutorizacao
    sqlSolicita = "Select id from tbAutorizacao order by id desc"
    rsSolicita.Open sqlSolicita, cnBanco, adOpenForwardOnly
    
    vPDO = rsSolicita.Fields(0)
    
    rsSolicita.Close
    Set rsSolicita = Nothing
    
    'Grava o numero da solicitação na tabela tbcolaboradores no campo
    'autorizacao

    sqlCandColab = "Update tbColaboradores set autorizacao = '" & vPDO & "' Where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "'"
    rsCandColab.Open sqlCandColab, cnBanco
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

Public Function caculaTmpExp()
On Error GoTo Err
    Dim rsTmpExp As New ADODB.Recordset
    Dim SqlTmpEx As String
    Dim periodoEmMeses As Single
    Dim X As Integer, Y As Integer
    SqlTmpEx = "select a.cpf,a.nomecolaborador,b.codmatriz,d.codcargo,d.nomecargo,b.data from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf " & _
    "inner join tbmatriz as c on b.codmatriz=c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo where b.ativo = 'S' and a.tipo = 'colaborador'"
    rsTmpExp.Open SqlTmpEx, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsTmpExp.RecordCount <> 0 Then Principal.ProgressBar1.Max = rsTmpExp.RecordCount
    X = 0
    Legenda = "Aguarde, reavaliando experiência dos colaboradores..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    'While Not rsTmpExp.EOF
    'Do While Not rsTmpExp.EOF
    For Y = 1 To rsTmpExp.RecordCount
        Principal.ProgressBar1.Value = X
        periodoEmMeses = DateDiff("m", rsTmpExp.Fields(5), Now)
        If Val(periodoEmMeses) > 0 Then
            'registraExperiencia rsTmpExp.Fields(0), rsTmpExp.Fields(3), periodoEmMeses
        End If
        rsTmpExp.MoveNext
        X = X + 1
    Next
    'Loop
    'Wend
    Principal.ProgressBar1.Value = 0
    Principal.StatusBar1.Panels(3).Text = "Grupo: " & GrupoUsu
    Set rsVExp = Nothing
    rsTmpExp.Close
    Set rsTmpExp = Nothing
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

Public Sub ajusta_LV()
On Error GoTo Err
    'ESSA ROTINA RESTAURA AS CONFIGURAÇÕES DE POSICIONAMENTO E TAMANHO DAS COLUNAS
    'DEFINIDAS PELO USUÁRIO
    Dim rsConfColunas As New ADODB.Recordset
    Dim SqlonfColunas As String
    Dim X As Integer, Y As Integer
    SqlonfColunas = "select * from tbConfLV where codcoligada = '" & vCodcoligada & "' and nmusuario = '" & NomUsu & "' and idmodulo = '" & apontaLV & "' order by codcoligada,nmusuario,idmodulo,posicao"
    rsConfColunas.Open SqlonfColunas, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsConfColunas.EOF Then
        Y = rsConfColunas.RecordCount
        While Not rsConfColunas.EOF
            MeuLV.ListView1.ColumnHeaders.Item(Val(rsConfColunas(2))).Position = Val(rsConfColunas(3))
            MeuLV.ListView1.ColumnHeaders.Item(Val(rsConfColunas(2))).Width = Val(rsConfColunas(4))
            rsConfColunas.MoveNext
        Wend
    End If
    MeuLV.ListView1.Refresh
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

Public Sub ajusta_LVTeste(vListView As Listview)
On Error GoTo Err
    'ESSA ROTINA RESTAURA AS CONFIGURAÇÕES DE POSICIONAMENTO E TAMANHO DAS COLUNAS
    'DEFINIDAS PELO USUÁRIO
    Dim rsConfColunas As New ADODB.Recordset
    Dim SqlonfColunas As String
    Dim X As Integer, Y As Integer
    SqlonfColunas = "select * from tbConfLV where codcoligada = '" & vCodcoligada & "' and nmusuario = '" & NomUsu & "' and idmodulo = '" & apontaLV & "' order by codcoligada,nmusuario,idmodulo,posicao"
    rsConfColunas.Open SqlonfColunas, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsConfColunas.EOF Then
        Y = rsConfColunas.RecordCount
        While Not rsConfColunas.EOF
            vListView.ColumnHeaders.Item(Val(rsConfColunas(2))).Position = Val(rsConfColunas(3))
            vListView.ColumnHeaders.Item(Val(rsConfColunas(2))).Width = Val(rsConfColunas(4))
            rsConfColunas.MoveNext
        Wend
    End If
    vListView.Refresh
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Sub


Public Sub GravarConfLV()
On Error GoTo Err
    'ESSA ROTINA RESTAURA AS CONFIGURAÇÕES DE POSICIONAMENTO E TAMANHO DAS COLUNAS
    'DEFINIDAS PELO USUÁRIO.
    'A TABELA TBCONFLV ARMAZENA AS CONFIGURAÇÕES DE POSICIONAMENTO E TAMANHO DAS COLUNAS.
    
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim X As Integer, Y As Integer
10  cnBanco.BeginTrans
   
    SqlSalvar = "Delete from tbConfLV where codcoligada = '" & vCodcoligada & "' and nmusuario = '" & NomUsu & "' and idmodulo = '" & apontaLV & "'"
    rsSalvar.Open SqlSalvar, cnBanco
    
    SqlSalvar = "select * from tbConfLV"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    Y = MeuLV.ListView1.ColumnHeaders.Count
    For X = 1 To Y
        rsSalvar.AddNew
        rsSalvar.Fields(0) = NomUsu 'Nome do usuário
        rsSalvar.Fields(1) = apontaLV 'id do módulo
        rsSalvar.Fields(2) = MeuLV.ListView1.ColumnHeaders.Item(X).Index 'índice da coluna
        rsSalvar.Fields(3) = MeuLV.ListView1.ColumnHeaders.Item(X).Position 'posição da coluna
        rsSalvar.Fields(4) = MeuLV.ListView1.ColumnHeaders.Item(X).Width 'largura da coluna
        rsSalvar.Fields(5) = vCodcoligada 'código da coligada
    Next
    rsSalvar.Update
    cnBanco.CommitTrans
    rsSalvar.Close
    Set rsSalvar = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    End If
End Sub

Public Function Enter(Key As Integer) As Integer
   If Key = 13 Then
       Enter = 0
   Else
       Enter = Key
   End If
End Function

Public Function NovoCodigo()
On Error GoTo Err
    Dim rsNovoCodigo As New ADODB.Recordset
    Dim sqlNovoCodigo As String
    sqlNovoCodigo = "select * from tbconfiguracoes"
    rsNovoCodigo.Open sqlNovoCodigo, cnBanco, adOpenKeyset, adLockReadOnly
    If QualForm = "novafo" Then
        If rsNovoCodigo.RecordCount > 0 Then
            If Val(rsNovoCodigo.Fields(0)) <> 0 Then NovoCodigo = Val(rsNovoCodigo.Fields(0)) Else NovoCodigo = Val(rsNovoCodigo.Fields(0)) + 1
        Else
            NovoCodigo = 1
        End If
    ElseIf QualForm = "novafce" Then
        If rsNovoCodigo.RecordCount > 0 Then
            If Val(rsNovoCodigo.Fields(1)) <> 0 Then NovoCodigo = Val(rsNovoCodigo.Fields(1)) Else NovoCodigo = Val(rsNovoCodigo.Fields(1)) + 1
        Else
            NovoCodigo = 1
        End If
    ElseIf QualForm = "novalm" Then
        If rsNovoCodigo.RecordCount > 0 Then
            If Val(rsNovoCodigo.Fields(2)) <> 0 Then NovoCodigo = Val(rsNovoCodigo.Fields(2)) Else NovoCodigo = Val(rsNovoCodigo.Fields(2)) + 1
        Else
            NovoCodigo = 1
        End If
    ElseIf QualForm = "novaos" Then
        If rsNovoCodigo.RecordCount > 0 Then
            If Val(rsNovoCodigo.Fields(3)) <> 0 Then NovoCodigo = Val(rsNovoCodigo.Fields(3)) Else NovoCodigo = Val(rsNovoCodigo.Fields(3)) + 1
        Else
            NovoCodigo = 1
        End If
    ElseIf QualForm = "novorel" Then
        If rsNovoCodigo.RecordCount > 0 Then
            If Val(rsNovoCodigo.Fields(4)) <> 0 Then NovoCodigo = Val(rsNovoCodigo.Fields(4)) Else NovoCodigo = Val(rsNovoCodigo.Fields(4)) + 1
        Else
            NovoCodigo = 1
        End If
    End If
    rsNovoCodigo.Close
    Set rsNovoCodigo = Nothing
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

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'A Função abaixo gera código para qualquer Listview
Public Function GeraCodigoLV(LV As Listview)
    If LV.ListItems.Count > 0 Then
        Dim X As Integer
        X = 1
        LV.Sorted = True
        LV.SortKey = 0
        LV.SortOrder = lvwDescending
        LV.ListItems.Item(X).Selected = True
        GeraCodigoLV = Val(LV.ListItems.Item(X)) + 1
    
        If apontaLV = 9 And LV.Name = "ListView1" Then
            LV.SortKey = 11
        End If
        LV.SortOrder = lvwAscending
        Exit Function
    Else
        GeraCodigoLV = 1
    End If
End Function

'A Função abaixo gera código para qualquer Tabela
Public Function GeraCodigoTB(vTabela As String, vCampo As String, vCampo2 As String, vText As String)
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    If vCampo2 = "" Or vCampo2 = "TERC" Then
        SqlGera = "Select top 1 * from " & vTabela & " order by " & vCampo & " Desc"
    'ElseIf vCampo2 = "TERC" Then
    '    SqlGera = "Select top 1 * from " & vTabela & " order by " & Mid$(vCampo, 5, 20) & " Desc"
    Else
        SqlGera = "Select top 1 * from " & vTabela & " where " & vCampo2 & "=" & Val(vText) & " order by " & vCampo & " Desc"
    End If
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraCodigo.RecordCount > 0 Then
        If vCampo2 = "" Then
            GeraCodigoTB = rsGeraCodigo.Fields(0) + 1
        Else
            GeraCodigoTB = Val(Mid$(rsGeraCodigo.Fields(0), 6, 20)) + 1
        End If
    Else
        GeraCodigoTB = 1
    End If
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
End Function

'A Função abaixo chama grid para quaisquer: textbox e tabela
Public Function ChamaGrid(vTabela As String, vCampo As String, vTxt As TextBox, vForm As Form, vPesq1 As String, vPesq2 As String)
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    If vTabela = "tbGrupoClass" Then
        Sqlp = "Select " & vPesq1 & "," & vPesq2 & " from " & vTabela & " where idprd ='" & frmFormulaCC.txtformula(0) & "' order by " & vCampo & ""
    ElseIf vTabela = "CORPORERM.dbo.GCCUSTO" Then
        'Somente será exibido CC que possuirem fórmula
        Sqlp = "select a.CODREDUZIDO,a.NOME from " & vBancoTotvs & ".dbo.GCCUSTO as a left join " & sDatabaseName & ".dbo.tbFormula as b on a.CODREDUZIDO = b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS " & _
            "Where a.ATIVO  = 'T' and b.nmform is not null and a.codcoligada = '" & vCodcoligada & "' group by a.ID,a.CODREDUZIDO,a.NOME order by a.CODREDUZIDO"
'            "Where a.ATIVO  = 'T' and substring(a.CODREDUZIDO,1,4) = '3000' and b.nmform is not null or ativo = 'T' and substring(a.CODREDUZIDO,1,12) = '7000.7103.SC' and b.nmform is not null or ativo = 'T' and substring(a.CODREDUZIDO,1,12) = '5000.5102.SC' and b.nmform is not null or a.ATIVO  = 'T' and substring(a.CODREDUZIDO,1,4) = '4000' and b.nmform is not null group by a.ID,a.CODREDUZIDO,a.NOME order by a.CODREDUZIDO"
    Else
        Sqlp = "Select " & vPesq1 & "," & vPesq2 & " from " & vTabela & "  where codcoligada = " & vCodcoligada & " order by " & vCampo & ""
    End If
    procnom = vCampo
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa"
    'Pesquisa = vForm.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Function
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find vCampo & "=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            vTxt.Text = rsLocal.Fields(0)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
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

'A Função abaixo é referente a ALTERAÇÃO de dados de qualquer ListView com até 15 colunas
Public Function AlteraLV(LV As Listview, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
On Error GoTo Err
    Dim Y As Integer, X As Integer, Z As Integer
    Dim vRaptor(15) As String
    For X = LBound(vRaptor) To UBound(vRaptor)
        vRaptor(X) = ""
    Next X
    Y = LV.ListItems.Count
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then
            vponteiro = X
            Exit For
        End If
    Next
    
    'SOMENTE PARA ZEUS--
    If apontaLV = 9 Then
        '1º VERIFICA SE A DATA PREVISTA ESTA VAZIA
        If LV.SelectedItem.ListSubItems(5).Text <> "" And LV.SelectedItem.ListSubItems(5).Text <> "-" Then
            '2º VERIFICA SE A SEMANA ATUAL É MAIOR OU IGUAL A SEMANA PROGRAMADA
            If IsDate(LV.SelectedItem.ListSubItems(5).Text) Then
                If DatePart("ww", (Date), vbMonday, vbFirstFourDays) >= DatePart("ww", CDate(LV.SelectedItem.ListSubItems(5).Text), vbMonday, vbFirstFourDays) Then
                    'bloqueiaEdicaoMP False
                    chamaForm.SkinLabel20.Visible = True
                    chamaForm.SkinLabel20.Caption = "O período para alteração dos dados dessa operação expirou"
                    'mobjMsg.Abrir "O período para alteração ds dados dessa operação expirou", Ok, critico, "Atenção"
                    'Exit Function
                Else
                    'bloqueiaEdicaoMP True
                    chamaForm.SkinLabel20.Visible = False
                    chamaForm.SkinLabel20.Caption = "Programação não pode ser alterada. Já está sendo apropriada"
                End If
            End If
        End If
    End If
    '-------------------
    
    If LV.ListItems.Count > 0 Then
        For Z = 1 To LV.ColumnHeaders.Count
            If Z = 1 Then
                vRaptor(Z) = LV.ListItems.Item(X)
            Else
                vRaptor(Z) = LV.SelectedItem.ListSubItems.Item(Z - 1)
            End If
        Next
    End If
    If vRaptor(1) <> "" Then vCP01.Text = vRaptor(1)
    If vRaptor(2) <> "" Then vCP02.Text = vRaptor(2)
    If vRaptor(3) <> "" Then vCP03.Text = vRaptor(3) 'Else vCP03.Text = ""
    If vRaptor(4) <> "" Then vCP04.Text = vRaptor(4)
    If vRaptor(5) <> "" Then vCP05.Text = vRaptor(5)
    If vRaptor(6) <> "" Then vCP06.Text = vRaptor(6)
    If vRaptor(7) <> "" Then vCP07.Text = vRaptor(7)
    If vRaptor(8) <> "" Then vCP08.Text = vRaptor(8)
    If vRaptor(9) <> "" Then vCP09.Text = vRaptor(9)
    If vRaptor(10) <> "" Then vCP10.Text = vRaptor(10) 'Else vCP10.Text = ""
    If vRaptor(11) <> "" Then vCP11.Text = vRaptor(11) 'Else vCP11.Text = ""
    If vRaptor(12) <> "" Then vCP12.Text = vRaptor(12) 'Else vCP12.Text = ""
    If vRaptor(13) <> "" Then vCP13.Text = vRaptor(13) 'Else vCP13.Text = ""
    If vRaptor(14) <> "" Then vCP14.Text = vRaptor(14) 'Else vCP14.Text = ""
    If vRaptor(15) <> "" Then vCP15.Text = vRaptor(15) 'Else vCP15.Text = ""
    Exit Function
Err:
    Resume Next
End Function

Private Sub bloqueiaEdicaoMP(vTipo As Boolean)

End Sub

'A Função abaixo é referente a INCLUSÃO de dados de qualquer ListView com até 10 colunas
Public Function IncluirLV(LV As Listview, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
    On Error Resume Next
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, Z As Integer
    Dim vRaptor(15) As TextBox
    
    Set vRaptor(1) = vCP01
    Set vRaptor(2) = vCP02
    Set vRaptor(3) = vCP03
    Set vRaptor(4) = vCP04
    Set vRaptor(5) = vCP05
    Set vRaptor(6) = vCP06
    Set vRaptor(7) = vCP07
    Set vRaptor(8) = vCP08
    Set vRaptor(9) = vCP09
    Set vRaptor(10) = vCP10
    Set vRaptor(11) = vCP11
    Set vRaptor(12) = vCP12
    Set vRaptor(13) = vCP13
    Set vRaptor(14) = vCP14
    Set vRaptor(15) = vCP15
    Y = LV.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
               If LV.ListItems.Item(X) = vRaptor(1) Then
                For Z = 1 To LV.ColumnHeaders.Count
                    If Z = 1 Then
                        If vRaptor(Z) <> "" Then vRaptor(Z) = LV.ListItems.Item(X)
                    Else
                        If vRaptor(Z) <> "" Then LV.SelectedItem.ListSubItems.Item(Z - 1) = vRaptor(Z)
                    End If
                Next
                Y = LV.ListItems.Count
                IncluirLV = True
                Exit Function
            End If
        Next
        If Formulario = "Métodos & Processos" And LV.Name = "ListView1" Then
            If separaDesLv(chamaForm.Text1.Text) = False Then
                IncluirLV = True
                Exit Function
            Else
                IncluirLV = True
            End If
        End If
        Set ItemLst = LV.ListItems.Add(, , vRaptor(1))
        Y = LV.ListItems.Count
    Else
        If Formulario = "Métodos & Processos" And LV.Name = "ListView1" Then
'        If chamaForm.Name = "frmMPCompleto" And LV.Name = "ListView1" Then
            If separaDesLv(chamaForm.Text1.Text) = False Then
                IncluirLV = False
                Exit Function
            Else
                IncluirLV = True
            End If
        End If
        Set ItemLst = LV.ListItems.Add(, , vRaptor(1))
        Y = LV.ListItems.Count
    End If
    For Z = 2 To LV.ColumnHeaders.Count
        If vRaptor(Z) <> "" Then ItemLst.SubItems(Z - 1) = vRaptor(Z)
    Next
    If vRaptor(2).Visible = True And vRaptor(2).Enabled = True Then
        vRaptor(2).SetFocus
    Else
        vRaptor(3).SetFocus
    End If
End Function

'Essa rotina serve para verificar se o item/c.custo que esta sendo inserido no ListView1
'Ja está em uma outra OS
Private Function separaDesLv(vTxtForm As String)
On Error GoTo Err
    separaDesLv = True
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    Dim vCodLM As String, vCodSeq As String
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If Len(RECEBE) = 5 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 3)
            Else
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 4)
            End If
            SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
            "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoTotvs & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
            "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoTotvs & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
            "inner join tbProjetos as g on g.codprojeto = d.codprojeto where a.fce = '" & Val(chamaForm.txtformula(12)) & "' and a.projeto = '" & chamaForm.txtformula(13).Text & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "' and a.idcc = '" & chamaForm.txtformula(0) & "' and a.idoperacao ='" & chamaForm.Combo1 & "'"
            rsTransf.Open SqlTransf, cnBanco, adOpenKeyset, adLockReadOnly
            If rsTransf.RecordCount > 0 Then
                mobjMsg.Abrir "Desenho: " & rsTransf.Fields(8) & vbCrLf & _
                              "Posição: " & rsTransf.Fields(11) & vbCrLf & _
                              "Item:" & rsTransf.Fields(12) & vbCrLf & _
                              "C.Custo:" & rsTransf.Fields(6) & vbCrLf & _
                              "Registrado na OS:" & Format(rsTransf.Fields(0), "000000000") & " - Programação: " & Format(rsTransf.Fields(7), "000000"), Ok, critico, "Atenção"
                separaDesLv = False
                rsTransf.Close
                Set rsTransf = Nothing
                Exit Function
            End If
            rsTransf.Close
            Set rsTransf = Nothing
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If Len(RECEBE) = 5 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 3)
        Else
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 4)
        End If
        SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
        "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoTotvs & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
        "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoTotvs & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
        "inner join tbProjetos as g on g.codprojeto = d.codprojeto where a.fce = '" & Val(chamaForm.txtformula(12)) & "' and a.projeto = '" & chamaForm.txtformula(13).Text & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "' and a.idcc = '" & chamaForm.txtformula(0) & "' and a.idoperacao ='" & chamaForm.Combo1 & "'"
        rsTransf.Open SqlTransf, cnBanco, adOpenKeyset, adLockReadOnly
        If rsTransf.RecordCount > 0 Then
            mobjMsg.Abrir "Desenho: " & rsTransf.Fields(8) & vbCrLf & _
                          "Posição: " & rsTransf.Fields(11) & vbCrLf & _
                          "Item:" & rsTransf.Fields(12) & vbCrLf & _
                          "C.Custo:" & rsTransf.Fields(6) & vbCrLf & _
                          "Registrado na OS:" & Format(rsTransf.Fields(0), "000000000") & " - Programação: " & Format(rsTransf.Fields(7), "000000"), Ok, critico, "Atenção"
            separaDesLv = False
            rsTransf.Close
            Set rsTransf = Nothing
            Exit Function
        End If
        rsTransf.Close
        Set rsTransf = Nothing
    End If
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

'A Função abaixo é referente a VALIDAÇÃO de dados de qualquer ListView com até 10 colunas
Public Function ValidaCampos(LV As Listview, vTxt1 As TextBox, vTxt2 As TextBox, vTxt3 As TextBox, vTxt4 As TextBox, vTxt5 As TextBox, vTxt6 As TextBox, vTxt7 As TextBox, vTxt8 As TextBox, vTxt9 As TextBox, vTxt10 As TextBox, vTxt11 As TextBox, vTxt12 As TextBox, vTxt13 As TextBox, vTxt14 As TextBox, vTxt15 As TextBox)
    'On Error Resume Next
    ValidaCampos = False
    Dim X As Integer
    Dim vMatrix(15) As TextBox
    'For X = LBound(vMatrix) To UBound(vMatrix)
    '    vMatrix(X) = ""
    'Next X
    Set vMatrix(1) = vTxt1
    Set vMatrix(2) = vTxt2
    Set vMatrix(3) = vTxt3
    Set vMatrix(4) = vTxt4
    Set vMatrix(5) = vTxt5
    Set vMatrix(6) = vTxt6
    Set vMatrix(7) = vTxt7
    Set vMatrix(8) = vTxt8
    Set vMatrix(9) = vTxt9
    Set vMatrix(10) = vTxt10
    Set vMatrix(11) = vTxt11
    Set vMatrix(12) = vTxt12
    Set vMatrix(13) = vTxt13
    Set vMatrix(14) = vTxt14
    Set vMatrix(15) = vTxt15
    For X = 1 To LV.ColumnHeaders.Count
        If vMatrix(X) = "" Then
            mobjMsg.Abrir "Favor informar o campo: " & vMatrix(X).Tag, Ok, informacao, "Atenção"
'            vMatrix(X).SetFocus
            Exit Function
        End If
    Next
    ValidaCampos = True
End Function

'A Função abaixo é referente a LIMPA DADOS de qualquer TetBox de ListView com até 10 colunas
Public Function LimpaControles(vTxt1 As TextBox, vTxt2 As TextBox, vTxt3 As TextBox, vTxt4 As TextBox, vTxt5 As TextBox, vTxt6 As TextBox, vTxt7 As TextBox, vTxt8 As TextBox, vTxt9 As TextBox, vTxt10 As TextBox)
    Dim X As Integer
    Dim vMatrix(10) As TextBox
    'ReDim vMatrix(10) 'Limpar Array
    Set vMatrix(1) = vTxt1
    Set vMatrix(2) = vTxt2
    Set vMatrix(3) = vTxt3
    Set vMatrix(4) = vTxt4
    Set vMatrix(5) = vTxt5
    Set vMatrix(6) = vTxt6
    Set vMatrix(7) = vTxt7
    Set vMatrix(8) = vTxt8
    Set vMatrix(9) = vTxt9
    Set vMatrix(10) = vTxt10
    For X = 1 To 10
       vMatrix(X).Text = ""
    Next
End Function

'A Função abaixo LIMPA DADOS de qualquer ListView
Public Function LimpaLV(LV As Listview)
    LV.ListItems.Clear
End Function

'A Função abaixo preenche dados da TXT2 baseado no dado informado na TXT1
Public Function CarregaTxt(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vCampo2 As String, vTipoCampo2 As String, vVar1 As TextBox, vVar2 As TextBox, vPosicao1 As Integer, vPosicao2 As Integer, vRetorno1 As TextBox, vTipoRetorno1 As String, vRetorno2 As TextBox, vQualQuery As String)
On Error GoTo Err
    'vTabela       = Nome da tabela a qual será realizada a pesquisa da Query
    'vCampo1       = Nome do campo da 1ª condição de pesquisa da Query
    'vTipoCampo1   = Tipo do 1º campo de pesquisa da Query
    'vCampo2       = Nome do campo da 2ª condição de pesquisa da Query
    'vTipoCampo2   = Tipo do 2º campo de pesquisa da Query
    'vVar1         = Nome do 1º TextBox que contem o valor que será pesquisado no 1ª campo de pesquisa da Query
    'vVar2         = Nome do 2º TextBox que contem o valor que será pesquisado no 2ª campo de pesquisa da Query
    'vPosicao1     = Posição do 1º campo que a Query irá retorna na consulta
    'vPosicao2     = Posição do 2º campo que a Query irá retorna na consulta
    'vRetorno1     = Nome 1º TextBox que irá receber o valor do 1º campo de retorno da Query
    'vTipoRetorno1 = Tipo do 1º campo de retorno (S/I - String ou Integer)
    'vRetorno2     = Nome 2º TextBox que irá receber o valor do 2º campo de retorno da Query
    'vQualQuery    = Qual query a função irá usar (1 - Uma das Querys que estão não função / 2 - Query fornecida pelo desenvolvedor  )
    
    Dim X As Integer
    Dim rsCarregaTxt As New ADODB.Recordset
    Dim sqlCarregaTxt As String
    
    If vQualQuery = 1 Then
        If vCampo2 = "" Then
            'Testa se vCampo1 é String ou Integer
            If vTipoCampo1 = "S" Then
                sqlCarregaTxt = "Select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "' order by '" & vCampo1 & "'"
            Else
                sqlCarregaTxt = "Select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "' order by '" & vCampo1 & "'"
            End If
        Else
            'Testa se vCampo1 e vCampo2 são String ou Integer
            If vTipoCampo1 = "S" And vTipoCampo2 = "S" Then
                sqlCarregaTxt = "Select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "' and " & vCampo2 & " = '" & vVar2 & "' order by '" & vCampo1 & "','" & vCampo2 & "'"
            ElseIf vTipoCampo1 = "S" And vTipoCampo2 = "I" Then
                sqlCarregaTxt = "Select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "' and " & vCampo2 & " = '" & Val(vVar2) & "' order by '" & vCampo1 & "','" & vCampo2 & "'"
            ElseIf vTipoCampo1 = "I" And vTipoCampo2 = "S" Then
                sqlCarregaTxt = "Select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "' and " & vCampo2 & " = '" & vVar2 & "' order by '" & vCampo1 & "','" & vCampo2 & "'"
            ElseIf vTipoCampo1 = "I" And vTipoCampo2 = "I" Then
                sqlCarregaTxt = "Select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "' and " & vCampo2 & " = '" & Val(vVar2) & "' order by '" & vCampo1 & "','" & vCampo2 & "'"
            End If
        End If
    Else
        sqlCarregaTxt = Sqlp
    End If
    rsCarregaTxt.Open sqlCarregaTxt, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCarregaTxt.EOF Then rsCarregaTxt.MoveFirst
    If rsCarregaTxt.EOF Then
        'mobjMsg.Abrir "Dado não cadastrado", Ok, informacao, "Atenção"
    Else
        If vTipoRetorno1 = "S" Then
            vRetorno1.Text = rsCarregaTxt.Fields(vPosicao1)
        Else
            vRetorno1.Text = Format(rsCarregaTxt.Fields(vPosicao1), "00")
        End If
        vRetorno2.Text = rsCarregaTxt.Fields(vPosicao2)
    End If
    rsCarregaTxt.Close
    Set rsCarregaTxt = Nothing
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

'A Função abaixo grava dados de até 50 variáveis de um formulário em uma determinada tabela
'Se algum campo não for preciso gravar ou alterar os dados (identifique-o, mas nos dois parâmetros deixe apenas as aspas sem nada)
Public Function GravaDados(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vVar1 As TextBox, vQtdCampos As Integer, vCampo2 As String, vTipoCampo2 As String, vVar2 As TextBox)
On Error GoTo Err
    'vTabela     = Nome da tabela a qual será realizada a pesquisa da Query
    'vCampo1     = Nome do campo da 1ª condição de pesquisa da Query
    'vTipoCampo1 = Tipo do 1º campo de pesquisa da Query
    'vVar1       = Nome do 1º TextBox que contem o valor que será pesquisado no 1ª campo de pesquisa da Query
    'vQtdCampos  =  Quantidade de variáveis que serão gravados na tabela
    'vCampo2     = Nome do campo da 2ª condição de pesquisa da Query
    'vTipoCampo2 = Tipo do 2º campo de pesquisa da Query
    'vVar2       = Nome do 2º TextBox que contem o valor que será pesquisado no 2ª campo de pesquisa da Query
    
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer, X As Integer
10  If vTransacaoAtiva = 0 Then cnBanco.BeginTrans
   
    If vTipoCampo1 = "I" Then
        If vCampo2 = "" Then
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "'"
        Else
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "' and " & vCampo2 & " = '" & Val(vVar2) & "'"
        End If
    Else
        SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "'"
    End If
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvar.EOF Then rsSalvar.AddNew
    For X = 0 To vQtdCampos - 1
        If vQualquerDado(X + 1, 2) = "S" Then
            rsSalvar.Fields(X) = vQualquerDado(X + 1, 1)
        ElseIf vQualquerDado(X + 1, 2) = "I" Then
            rsSalvar.Fields(X) = Val(vQualquerDado(X + 1, 1))
        ElseIf vQualquerDado(X + 1, 2) = "D" Then
            If vQualquerDado(X + 1, 1) <> "" Then
                rsSalvar.Fields(X) = CDate(vQualquerDado(X + 1, 1))
            End If
        End If
    Next
    rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    If vTransacaoAtiva = 0 Then cnBanco.CommitTrans
    'MsgBox "Os dados do FORMULARIO foram salvos com sucesso", vbInformation, "ProtótipoX"
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

'A função abaixo limpa resíduos de dados da matriz/Array vQualquerDado
Public Function limpaQualquerDado()
    Dim X As Integer, Y As Integer
'    For X = LBound(vQualquerDado) To UBound(vQualquerDado)
    For X = 1 To 50
        For Y = 1 To 30
            vQualquerDado(X, Y) = ""
        Next
    Next X
End Function

'A função abaixo Armazena todos os dados da ListView na Array vQualquerDado na posição correta
'a serem gravado na Tabela
'Se o dado a ser ordenado não estiver no Listview, deve-se informar o nome do componente.propriedade
'LV             = Nome do Listview com o qual a função irá trabalhar
'vPos0 ~ vPos10 = Posição das colunas e componentes os quais irão ser organizados/armazenados pela função

Public Function ordenaLVArray(LV As Listview, vPos0 As String, vPos1 As String, vPos2 As String, vPos3 As String, vPos4 As String, vPos5 As String, vPos6 As String, vPos7 As String, vPos8 As String, vPos9 As String, vPos10 As String, vPos11 As String, vPos12 As String, vPos13 As String, vPos14 As String, vPos15 As String)
On Error Resume Next
    Dim X As Integer, Y As Integer, Z As Integer
    Dim vMatrix(16) As String
    For X = LBound(vMatrix) To UBound(vMatrix)
        vMatrix(X) = ""
    Next X
    vMatrix(1) = vPos0
    vMatrix(2) = vPos1
    vMatrix(3) = vPos2
    vMatrix(4) = vPos3
    vMatrix(5) = vPos4
    vMatrix(6) = vPos5
    vMatrix(7) = vPos6
    vMatrix(8) = vPos7
    vMatrix(9) = vPos8
    vMatrix(10) = vPos9
    vMatrix(11) = vPos10
    vMatrix(12) = vPos11
    vMatrix(13) = vPos12
    vMatrix(14) = vPos13
    vMatrix(15) = vPos14
    vMatrix(16) = vPos15
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        For Z = 0 To LV.ColumnHeaders.Count
            If IsNumeric(vMatrix(Z + 1)) Then
                If vMatrix(Z + 1) = "0" Then
                    vQualquerDado(X, Z + 1) = LV.ListItems.Item(X)
                Else
                    'Se o valor da Listview for igual a "-" grava zero
                    If LV.SelectedItem.ListSubItems.Item(Val(vMatrix(Z + 1))) = "-" Then
                        vQualquerDado(X, Z + 1) = 0
                    Else
                        If LV.SelectedItem.ListSubItems.Item(Val(vMatrix(Z + 1))) = "" Then
                            vQualquerDado(X, Z + 1) = " "
                        Else
                            vQualquerDado(X, Z + 1) = LV.SelectedItem.ListSubItems.Item(Val(vMatrix(Z + 1)))
                        End If
                        
                        'vQualquerDado(X, Z + 1) = vMatrix(Z + 1)
                    End If
                End If
            Else
                vQualquerDado(X, Z + 1) = vMatrix(Z + 1)
            End If
        Next
    Next
End Function

'A Função abaixo grava dados em tabela de um determinado Listview de até 50 variáveis aramazenadas na sequencia
'corrreta, ordenado pela função chamada anterior a essa
Public Function GravaDadosLV(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vVar1 As TextBox)
On Error GoTo Err
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer, X As Integer
    
10  cnBanco.BeginTrans
    
    If vCampo1 = "" Then
        sqlDeletar = "Delete from " & vTabela
    Else
        If vTipoCampo1 = "I" Then
            sqlDeletar = "Delete from " & vTabela & " where " & vCampo1 & " ='" & Val(vVar1) & "'"
        Else
            sqlDeletar = "Delete from " & vTabela & " where " & vCampo1 & " ='" & vVar1 & "'"
        End If
    End If
    rsDeletar.Open sqlDeletar, cnBanco
      
    SqlSalvar = "select * from " & vTabela
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    X = 1
    Y = 1
    While vQualquerDado(X, Y) <> ""
        rsSalvar.AddNew
        'If Asc(vQualquerDado(X, Y)) = 160 Then vQualquerDado(X, Y) = ""
        While vQualquerDado(X, Y) <> "" 'Or vQualquerDado(X, Y) <> " "
            
            If apontaLV = 9 And Y = 10 Then 'Essa condição serve Somente para o zeus
                rsSalvar.Fields(9) = Mid$(vQualquerDado(X, Y), 1, 9)
                rsSalvar.Fields(15) = Mid$(vQualquerDado(X, Y), 11, 1)
            Else
                If vQualquerDado(X, Y) <> "-" Then rsSalvar.Fields(Y - 1) = vQualquerDado(X, Y)
            End If
            Y = Y + 1
            'If Asc(vQualquerDado(X, Y)) = 160 Then vQualquerDado(X, Y) = ""
        Wend
        X = X + 1
        Y = 1
    Wend
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    cnBanco.CommitTrans
    'MsgBox "Os dados do LISTVIEW foram salvos com sucesso", vbInformation, "ProtótipoX"
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        'Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbCritical, "Atenção"
        'cnBanco.RollbackTrans
        'Exit Function
        Resume Next
    End If
End Function

Public Function chamaSQL(vSql As String)
    Sqlp = ""
    Sqlp = vSql
End Function

Public Function Compoe_Listview(LV As Listview, vSqlCompoe As String, vZerosEsq As String)
On Error GoTo Err
    ' Declaração de variaveis
    Dim rsCompoe As New ADODB.Recordset
    If vZerosEsq <> "TOTVS" Then
        rsCompoe.Open vSqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    Else
        rsCompoe.Open vSqlCompoe, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    End If

    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsCompoe.EOF
        For X = 0 To LV.ColumnHeaders.Count - 1
            If X = 0 Then
                If rsCompoe.Fields(X).Type = adInteger Then
                    Set ItemLst = LV.ListItems.Add(, , Format(rsCompoe.Fields(X), vZerosEsq))
                Else
                    Set ItemLst = LV.ListItems.Add(, , rsCompoe.Fields(X))
                End If
            Else
                If chamaForm.Name = "frmMPCompleto" And LV.Name = "ListView1" And X = 1 Then
                    If rsCompoe.Fields(X) = 0 Then
                        ItemLst.SubItems(X) = "0"
                    Else
                        ItemLst.SubItems(X) = rsCompoe.Fields(X) '"" & Format(rsCompoe.Fields(X), "000000000")
                    End If
                Else
                    If rsCompoe.Fields(X) = 0 Then
                        ItemLst.SubItems(X) = "0"
                    Else
                        ItemLst.SubItems(X) = "" & rsCompoe.Fields(X)
                    End If
                End If
            End If
        Next
        rsCompoe.MoveNext
    Wend
    LV.Sorted = True
    If apontaLV = 9 And LV.Name = "Listview1" Then
        LV.SortKey = 11
    Else
        LV.SortKey = 0
    End If
    LV.SortOrder = lvwAscending
    rsCompoe.Close
    Set rsCompoe = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Function


Public Function Compoe_Listview2(LV As Listview, vSqlCompoe As String, vZerosEsq As String)
On Error GoTo Err
    ' Declaração de variaveis
    Dim rsCompoe As New ADODB.Recordset
    If vZerosEsq <> "TOTVS" Then
        rsCompoe.Open vSqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    Else
        rsCompoe.Open vSqlCompoe, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    End If

    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsCompoe.EOF
        For X = 0 To LV.ColumnHeaders.Count - 1
            If X = 0 Then
                If rsCompoe.Fields(X).Type = adInteger Then
                    Set ItemLst = LV.ListItems.Add(, , Format(rsCompoe.Fields(X), vZerosEsq))
                Else
                    Set ItemLst = LV.ListItems.Add(, , rsCompoe.Fields(X))
                End If
            Else
                    If rsCompoe.Fields(X) = 0 Then
                        ItemLst.SubItems(X) = "0"
                    Else
                        ItemLst.SubItems(X) = "" & rsCompoe.Fields(X)
                    End If
            End If
        Next
        rsCompoe.MoveNext
    Wend
    LV.Sorted = True
    If apontaLV = 9 And LV.Name = "Listview1" Then
        LV.SortKey = 11
    Else
        LV.SortKey = 0
    End If
    LV.SortOrder = lvwAscending
    rsCompoe.Close
    Set rsCompoe = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Function

Public Function mudaCorText(txt As TextBox)
    txt.BackColor = &HC0E0FF
End Function

Public Function voltaCorText(txt As TextBox)
    txt.BackColor = &HFFFFFF
End Function

Public Function mudaCorMask(msk As MaskEdBox)
    msk.BackColor = &HC0E0FF
End Function

Public Function voltaCorMask(msk As MaskEdBox)
    msk.BackColor = &HFFFFFF
End Function

Public Function selectColor(cor As TextBox, dlgCores As CommonDialog)
    With dlgCores
        .ShowColor
        cor.BackColor = dlgCores.Color
    End With
End Function


Public Function converteSemana(vSemanaAno As Integer, vDataAno As DTPicker, vAnoDaSemana As String)
    On Error GoTo Err
    Dim vDataConvertida As String 'Dia/Mes/Ano
    Dim vAno As String, vMes As String, vDia As String
    Dim vX As Integer, vY As Integer
    
    If vAnoDaSemana = "" Then
        vAno = DatePart("yyyy", Date, vbMonday, vbFirstFourDays)
    Else
        vAno = vAnoDaSemana
    End If
    
    For vX = 1 To 12
        vMes = Format(vX, "00")
        For vY = 1 To 31
            vDia = Format(vY, "00")
            vDataConvertida = Format(vDia & "/" & vMes & "/" & vAno, "dd/mm/yyyy")
            If IsDate(CDate(vDataConvertida)) Then
                If DatePart("ww", CDate(vDataConvertida), vbMonday, vbFirstFourDays) = vSemanaAno Then
                     vDataAno.Value = CDate(vDataConvertida)
                    Exit Function
                End If
            
            End If
        Next
    Next
Err:
    vX = vX + 1
    vY = 1
    vMes = Format(vX, "00")
    vDia = Format(vY, "00")
    vDataConvertida = Format(vDia & "/" & vMes & "/" & vAno, "dd/mm/yyyy")
    Resume Next
End Function

Public Function MarcaDesmarca(LV As Listview)
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems(X).Selected = True
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        Else
            LV.ListItems.Item(X).Checked = True
        End If
    Next
End Function

Public Function montaDadosVendas()
On Error GoTo Err
    Dim X As Integer, Y As Integer, j As Integer
    Dim rsVendas As New ADODB.Recordset
    Dim sqlVendas As String
    limpaQualquerDado
    strAno = Mid$(varGlobal, 1, 4)
    
    sqlVendas = "select a.id,a.numoc,a.descricao,a.quantidade,a.unqtd,a.peso,a.unpeso,a.valorsimp,a.pisperc,a.pisvalor,a.cofinsperc,a.cofinsvalor,a.icmsperc,a.icmsvalor,a.valorcimp,a.und,a.subtotal,a.ipiperc,a.ipivalor,a.total,a.bcalcicms,a.foreferente,a.condicaopag,a.adiantamento,a.adiantamentoCP from tbpedidos as a where fce ='" & strAno & "'"
    rsVendas.Open sqlVendas, cnBanco, adOpenKeyset, adLockReadOnly
    Y = rsVendas.RecordCount
    j = 1
    For X = 1 To Y
        If Not IsNull(rsVendas.Fields(1)) Then vQualquerDado(X, j) = rsVendas.Fields(1) 'Pedido
        If Not IsNull(rsVendas.Fields(2)) Then vQualquerDado(X, j + 1) = rsVendas.Fields(2) 'Descrição
        If Not IsNull(rsVendas.Fields(5)) Then vQualquerDado(X, j + 2) = rsVendas.Fields(5) 'Peso Total
        If Not IsNull(rsVendas.Fields(19)) Then vQualquerDado(X, j + 3) = rsVendas.Fields(19) 'Valor Total
        If Not IsNull(rsVendas.Fields(22)) Then vQualquerDado(X, j + 4) = rsVendas.Fields(22) 'Condição de pagamento
        
        If Not IsNull(rsVendas.Fields(23)) Then vQualquerDado(X, j + 5) = rsVendas.Fields(23) 'Adiantamento - %
        If Not IsNull(rsVendas.Fields(24)) Then vQualquerDado(X, j + 6) = rsVendas.Fields(24) 'Adiantamento - Condição de pagamento
        j = 1
        rsVendas.MoveNext
    Next
    rsVendas.Close
    Set rsVendas = Nothing

    sqlVendas = "select A.CODTMV,sum(d.VALORBAIXADO+D.VALORADIANTAMENTO) as VALORBAIXADO from " & vBancoTotvs & ".dbo.TMOV as a inner join " & vBancoTotvs & ".dbo.TCPG as  B ON A.CODCPG = B.CODCPG " & _
                "LEFT JOIN " & vBancoTotvs & ".dbo.TMOVCOMPL AS C ON A.IDMOV = C.IDMOV LEFT JOIN " & vBancoTotvs & ".dbo.FLAN AS D ON A.IDMOV = D.IDMOV INNER JOIN " & vBancoTotvs & ".dbo.TTB3 as E on A.CODTB3FAT = E.CODTB3FAT where a.CODTB3FAT = '" & strAno & "' and a.CODTMV='2.2.25' group by A.CODTMV "
    rsVendas.Open sqlVendas, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVendas.RecordCount > 0 Then vQualquerDado(20, 1) = rsVendas.Fields(1)
    rsVendas.Close
    Set rsVendas = Nothing


    sqlVendas = "select sum(d.VALORBAIXADO+D.VALORADIANTAMENTO) as VALORBAIXADO from " & vBancoTotvs & ".dbo.TMOV as a inner join " & vBancoTotvs & ".dbo.TCPG as  B ON A.CODCPG = B.CODCPG " & _
                "LEFT JOIN " & vBancoTotvs & ".dbo.TMOVCOMPL AS C ON A.IDMOV = C.IDMOV LEFT JOIN " & vBancoTotvs & ".dbo.FLAN AS D ON A.IDMOV = D.IDMOV INNER JOIN " & vBancoTotvs & ".dbo.TTB3 as E on A.CODTB3FAT = E.CODTB3FAT where a.CODTB3FAT = '" & strAno & "' and a.CODTMV in ('2.2.01','2.2.05','2.2.25')"
    rsVendas.Open sqlVendas, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVendas.RecordCount > 0 Then vQualquerDado(20, 2) = rsVendas.Fields(0)
    rsVendas.Close
    Set rsVendas = Nothing


    sqlVendas = "select A.CODTMV,sum(A.PESOBRUTO )from " & vBancoTotvs & ".dbo.TMOV as a where a.CODTB3FAT = '" & strAno & "' and a.CODTMV in ('2.2.01','2.2.05') and A.STATUS <> 'C' group by A.CODTMV "
    rsVendas.Open sqlVendas, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVendas.RecordCount > 0 Then vQualquerDado(20, 3) = rsVendas.Fields(1)
    rsVendas.Close
    Set rsVendas = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Function

Public Function CampoHora(obj As Object, Keyasc As Integer)
    If Not ((Keyasc >= Asc("0") And Keyasc <= Asc("9")) Or Keyasc = 8) Then
        Keyasc = 0
        Exit Function
    End If
    If Keyasc <> 8 Then
        If Len(obj.Text) = 2 Then
            obj.Text = obj.Text + ":"
            obj.SelStart = Len(obj.Text)
        End If
    End If
End Function

Public Function acertaTamanhoIcone()
    Dim H As Integer
    MeuLV.chameleonButton1.Height = 495
    MeuLV.cmdconsulta(0).Height = 615
    For H = 4 To 12
        MeuLV.cmdconsulta(H).Height = 615
    Next
End Function
