Attribute VB_Name = "RotinaGeral"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Biblioteca para manipulação do Regedit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
Public vTransacaoAtiva As Long

Public oConn As ADODB.Connection
Public sDatabaseName As String 'Utilizado para criar nova conexão com o banco na tela de splash
Public sServerName As String ' Utilizado para criar nova conexão com o banco na tela de splash
Public sUsuName As String 'Nome do usuário de conexão ao DB
Public sSenhaDB As String 'Senha de conexão ao DB
Public sSGBD As Integer 'Versão do SGBD
Public sEmailAvRec As String 'String que guarda endereços de e-mail que receberão notificações caso a avaliação de fornecimento esteja abaixo da média global
Public vFormatoDatetime As String

Public vDataFilter1 As String
Public vDataFilter2 As String
Public vDataExportMed As String

Public sLogoEmpresa As String ' Utilizado para guardar o caminho da logo da empresa
Public StatusTrei As String 'Verifica o status do treinamento
Public vLocalEstoque As String

Public rsVExp As New ADODB.Recordset
Public SqlVExp As String

Public cnBanco As ADODB.Connection
Public cnBancoSAP As ADODB.Connection
Public rsLocal As New ADODB.Recordset
Public Sqlp As String
'Public rsResumo As New ADODB.Recordset
Public MediaGlobal As Double
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
Public vQtdDisponivel As Integer, vQtdSolicitada As Integer 'armazena a quantidade disponvel de ferramentas

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
Public vSMTPAutentic As Integer  'grava endereço do servidor de SMTP
Public vSSL As Boolean 'grava endereço do servidor de SMTP
Public SerieEmpresa As String ' Serie da empresa
Public vMaiorIDMOV As Double 'Grava maior IDMOV depois da sincronizacao

Public vPorta As String 'grava porta do servidor de SMTP
Public vUsuEmail As String 'grava nome do usuario de autenticação
Public vSenhaEmail As String 'grava a senha do usuário de autenticação
Public vIntegra As String  'Para informar se o ZEUSH esta integrado a outro sistema
Public vDataDoBanco As Date 'Grava a data atual do Banco de dados
Public vDadosSAP(18) As String
Public vInicioAvOC As Date 'Aramazena a data de inicio das Avaliações das Ordens de Compra
Public vPeridoAvFornec As Date

Public vTipoAvaliador As String 'Armazena o tipo de colaborador que irá realizar a avaliação de recebimento das OCs (Recebimento/Setor técnico)
Public colheDados(17) As String 'Guarda dados de importação de colaboradores de arquivo TXT
Public FimAprop As String 'Verifica se o colaborador tem permissão de encerra apropriação de colaboradores que estão apropriando em alguma OS

Public mStream As ADODB.Stream 'Para gravar imagem no Banco SAP

Public vServerSAP As String  'Armazena nome do servidor SAP
Public vBancoSAP As String  'Armazena nome do banco SAP
Public vUsuBancoTovs As String  'Armazena usuario do banco SAP
Public vSenhaBancoSAP As String  'Armazena senha do banco SAP

Public chamaForm As Form

Public MeuLV As New frmPesqGeral
Public NomeColLV(25) As String
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

'***** VARIAVEIS DE INTEGRAÇÃO TOTVS RM SISTEMAS ******************************
Public vIDMov As Double
Public vNumeromov As String 'Variaveis de integração do RM
Public vSerie As String 'variável de integração do RM
Public vCodColigadaRM As Integer 'Variável de integração do RM
Public vCodVenRM As String 'Variável de integração do RM
Public vNomeVenRM As String 'Variável de integração do RM
Public vCodUsuarioRM As String 'Variavel de integralçao RM
Public vSequencialEstoque As Double 'Variavel de integração RM
Public vCodLocalEstoque As String 'variável de integração RM - Guarda o código do local de estoque
Public vLogin As String 'Variável de integração do RM
' VARIAVEIS QUE ARMAZENAM INFORMAÇÕES PARA FUNCIONAMENTO OFFLINE
Public vIntegraOffline As String  '
Public vBancoOffline As String 'Utilizado para criar nova conexão com o banco na tela de splash
Public vServerOffline As String ' Utilizado para criar nova conexão com o banco na tela de splash
Public vUsuBancoOffline As String 'Nome do usuário de conexão ao DB
Public vSenhaBancoOffline As String 'Senha de conexão ao DB
Public vCaminhoReg As String 'Armazena o caminho das pastas no regedit

'******************************************************************************

'Public CodUsu As String ' codigo do usuário q esta logado
'Public NomUsu As String ' Nome do usuario
'Public CapturaCodigo As String ' Codigo da Empresa e do Contato
'Public Legenda As String ' Informa o significado (F)Fone (F) Fax (C) Celular
Public procnom As String, procnom1 As String
Public strAno As String 'Usada no relatorio de programação de cursos/treinamentos anual
Public vQualquerDado(50, 30) As String
Public vTabela1 As String, vTabela2 As String, vTabela3 As String, vTabela4 As String, vTabela5 As String, vTabela6 As String, vTabela7 As String, vTabela8 As String, vTabela9 As String, vTabela10 As String, vTabela11 As String, vTabela12 As String, vTabela13 As String, vTabela14 As String, vTabela15 As String
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
On Error GoTo Err1
    Conexao = True
    If sServerName = "" Then GoTo Err1
    Set cnBanco = New ADODB.Connection
        
        'Conexão rede local (funcionando)
        'cnBanco.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=True;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
        cnBanco.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    
    frmSplash.Label5.Caption = "Conexão realizada com sucesso"
    Exit Function
Err1:
    frmSplash.Label5.Caption = "Falha na conexão: " & Err.Number & " - " & Err.Description
    Msgbox "Erro ao tentar acessar DB - Entre com as novas configurações do servidor ", vbCritical, "Atenção"
    Conexao = False
    Exit Function
End Function

'ABAIXO CONEXÃO COM O BANCO DE DADOS RM
Public Function ConexaoSAP()
'On Error GoTo Err1
    Set cnBancoSAP = New ADODB.Connection
    cnBancoSAP.Open "Provider=SQLOLEDB.1;Password=" & vSenhaBancoSAP & ";Persist Security Info=True;User ID=" & vUsuBancoTovs & ";Initial Catalog=" & vBancoSAP & ";Data Source=" & vServerSAP
    vIntegra = "S"
    'achaSecaoZEUSH
    'criaTrigger
    Exit Function
Err1:
    Msgbox "Erro de conexão com Banco SAP", vbCritical, "Atenção"
    'mobjMsg.Abrir "Erro de conexão com Banco SAP", Ok, critico, "Atenção"
    vIntegra = "N"
    Exit Function
End Function


Public Function ConexaoLdap()
    Dim sUser As String, sDN As String, sRoot As String
    sUser = "admin"
    sDN = "uid=" & sUser
    sRoot = "LDAP://10.10.10.29/phpldapadmin:389,dc=id"
    Dim oDS: Set oDS = GetObject("LDAP:")
    'On Error GoTo AuthError
    Dim oAuth: Set oAuth = oDS.OpenDSObject(sRoot, sDN, "049332id", &H200)
    'On Error GoTo 0
    Msgbox "Login Successful"
    Exit Function
AuthError:
    If Err.Number = -2147023570 Then
        Msgbox "Wrong Username or password !!!"
    End If
    On Error GoTo 0
End Function


Public Sub CompoeCombo(Combo As ComboBox, Tabela, campo, Campo1)
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
End Sub

Public Sub CompoeCombo1(Combo As ComboBox, Tabela, campo, Campo1)
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
End Sub

Public Sub CompoeCombo2(Combo As ComboBox, Tabela, campo, Campo1)
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
        sql = "Select " & Campo1 & " from " & Tabela & " Order By " & Campo1
    End If
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
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

Public Sub CompoeComboLVPesq(Combo As ComboBox, LV As ListView, vIndiceCombo As Integer, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    'If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            Combo.AddItem c
        Next
        Combo.Text = Combo.List(vIndiceCombo)
    'End If
End Sub

Public Sub CompoeComboCC(Combo As ComboBox)
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    sql = "select a.NOME from CORPORERM.dbo.GCCUSTO as a where a.ATIVO = 'T' and substring(a.nome,1,4) = '3000' or substring(a.nome,1,4) = '4000' or substring(a.nome,1,4) = '7000' or substring(a.nome,1,4) = '5000'"
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
End Sub

Public Sub CompoeComboSQL(Combo As ComboBox, vSql As String)
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
On Error GoTo Err1
    
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
    
    'CRIA BANCO IMRM
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDadosBanco(" & _
    "NomeServidor VARCHAR(50) NULL," & _
    "NomeBanco VARCHAR(50) NULL)"
    
    'TABELAS IMRM
'============================
    'CRIA TABELAS IMRM
    
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
    "inicioavaliacao DATETIME NULL," & _
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
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbFiltro(" & _
    "idfiltro INT NOT NULL IDENTITY," & _
    "usuario VARCHAR(50) NOT NULL," & _
    "modulo VARCHAR(100) NOT NULL," & _
    "tipofiltro VARCHAR(50) NOT NULL," & _
    "nomefiltro VARCHAR(50) NOT NULL," & _
    "query TEXT NOT NULL," & _
    "expressao VARCHAR(300) NULL," & _
    "padrao CHAR(1) NULL," & _
    "PRIMARY KEY (idfiltro))"
    
    
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
    "serie VARCHAR(3) NOT NULL," & _
    "PRIMARY KEY (codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfEmail(" & _
    "smtp VARCHAR(100) NULL," & _
    "usuario VARCHAR(50) NULL," & _
    "senha VARCHAR(30) NULL," & _
    "codcoligada INT NULL," & _
    "porta int NULL," & _
    "ssl int NULL," & _
    "smtpautentic int NULL)"

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
    "nome VARCHAR(100) NOT NULL," & _
    "codven VARCHAR(16) NULL," & _
    "nomeven VARCHAR(100) NULL," & _
    "email VARCHAR(50) NULL," & _
    "codgrupo NUMERIC NULL," & _
    "altlogin NUMERIC NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "codcoligadatotvs VARCHAR(100) NULL," & _
    "codusuarioTOTVS VARCHAR(50) NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbGrupo(" & _
    "codigo NUMERIC NOT NULL," & _
    "descricao VARCHAR(50) NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"
   
    
'============================
    'CRIA TABELAS ESPECIFICAS DO SISTEMA


    'Tabela de gerenciamento das medições de Terceiros
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMedicoesTerceiro(" & _
    "id INT NOT NULL IDENTITY," & _
    "codigo VARCHAR(20) NOT NULL," & _
    "idmovintegracao INT NOT NULL," & _
    "idmovnf INT NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "status INT NOT NULL," & _
    "usercadastro VARCHAR(10) NOT NULL," & _
    "dtcadastro DATETIME NOT NULL," & _
    "dtexport DATETIME NOT NULL," & _
    "PRIMARY KEY (codigo))"
    
    
    'Tabela de gerenciamento das medições de PJ
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMedicoesPJ(" & _
    "id INT NOT NULL IDENTITY," & _
    "codigo INT NOT NULL," & _
    "idmovintegracao INT NOT NULL," & _
    "idmovnf INT NULL," & _
    "observacao VARCHAR(300) NULL," & _
    "status INT NOT NULL," & _
    "usercadastro VARCHAR(10) NOT NULL," & _
    "dtcadastro DATETIME NOT NULL," & _
    "dtexport DATETIME NOT NULL," & _
    "PRIMARY KEY (codigo))"
    
    'Tabela que gera NUMEROMOV
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAvisos(" & _
    "id INT NOT NULL IDENTITY," & _
    "idmedicao VARCHAR(20) NOT NULL," & _
    "status INT NULL," & _
    "PRIMARY KEY (idmedicao))"
    
    'Tabela que gera NUMEROMOV
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMov(" & _
    "numeromov VARCHAR(35) NOT NULL," & _
    "serie VARCHAR(8) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (numeromov,serie,codcoligada))"
    
'============================
    
    'ABAIXO: CRIA CONFIGURAÇÃO PARA USUÁRIO ADMINISTRADOR
    oConn.Close
    
    vCodcoligada = 1 'Primeiro cadastro de coligada
    
    oConn.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=True;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName

    SqlSenha = "Insert into tbSenha(usuario,senha,codigo,codcoligada) Values('adm','123',1,'" & vCodcoligada & "');"
    rsSenha.Open SqlSenha, oConn
    
    SqlUsuario = "Insert into tbUsuarios(codigo,nome,codgrupo,ativo,codcoligada) Values(1,'Administrador do sistema',1,'S','" & vCodcoligada & "');"
    rsUsuario.Open SqlUsuario, oConn
    
    SqlGrupo = "Insert into tbGrupo(codigo,descricao,ativo,codcoligada) Values(1,'Administradores','S','" & vCodcoligada & "');"
    rsGrupo.Open SqlGrupo, oConn
    
    SqlConfGrupo = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'01','TAB','Cadastros','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'01','CAT','Primários','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'02','CAT','Secundários','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0101','BUT','','S','" & vCodcoligada & "',1);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0102','BUT','','S','" & vCodcoligada & "',2);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0103','BUT','','S','" & vCodcoligada & "',3);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0104','BUT','','S','" & vCodcoligada & "',4);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0205','BUT','','S','" & vCodcoligada & "',5);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0206','BUT','','S','" & vCodcoligada & "',6);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0207','BUT','','S','" & vCodcoligada & "',7);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,1,'0208','BUT','','S','" & vCodcoligada & "',8);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,2,'02','TAB','Movimentações','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,2,'11','CAT','Gestaõ de Movimentações','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,2,'1111','BUT','Emprestimos/Devoluções','S','" & vCodcoligada & "',9);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'06','TAB','Configurações','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'51','CAT','Parametrizações','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'52','CAT','Aparência','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5151','BUT','Sistema','S','" & vCodcoligada & "',17);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5152','BUT','Grupos','S','" & vCodcoligada & "',18);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5153','BUT','Usuários','S','" & vCodcoligada & "',19);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5254','BUT','Menu','S','" & vCodcoligada & "',20);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5255','BUT','Skin','S','" & vCodcoligada & "',21);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,6,'5256','BUT','Fundo','S','" & vCodcoligada & "',22);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'07','TAB','Sobre','S','" & vCodcoligada & "',0);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'61','CAT','Sobre','S','" & vCodcoligada & "',0);" & _
                   "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'6161','BUT','Sobre IMRM','S','" & vCodcoligada & "',23);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'6162','BUT','Ajuda do IMRM','S','" & vCodcoligada & "',24);"
    
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
       
    Msgbox "Tabelas criadas com sucesso", vbInformation, "IMRM"
    Exit Function
Err1:
    'Msgbox "(ADO) Erro ao criar Tabela de dados: " & vbCrLf & Err.Number & " - Tabela já Existe - " & Err.Description, 16, "Mensagem de erro"
    Resume Next
    'Exit Function
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
    If Formulario = "Empréstimo/Devolução" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Empréstimos"
        frmFiltro.Combo1.List(1) = "Devoluções"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Empréstimos"
    End If
    If Formulario = "Medição PJ/Mensal" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        'frmFiltro.Combo1.List(2) = "Todos"
        'frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Critérios de Avaliação de Fornecimento" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Fornecedores" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Grupo de Critérios de Avaliação de Fornecedores" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Recebimento de Pedido de Compra" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
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
    If Formulario = "LM's - Listas de Materiais" Then
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
    If Formulario = "Notas das Avaliações dos Fornecedores" Then
        'TiPo = False
        'frmFiltro.Combo1.List(0) = "Ativos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Usuários" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Grupos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
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
    If Formulario = "Método & Processo" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Planejamento"
        frmFiltro.Combo1.List(1) = "Produção"
        frmFiltro.Combo1.List(2) = "Expedição"
        frmFiltro.Combo1.List(3) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Relatórios de Inspeção" Then
        'TiPo = False
'        frmFiltro.Combo1.List(0) = "Avaliados"
'        frmFiltro.Combo1.List(1) = "Não Avaliados"
        frmFiltro.Combo1.List(0) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
    If Formulario = "Recebimento de Ordem de Compra" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Todos"
        'frmFiltro.Combo1.List(1) = "Não ativos"
        'frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Todos"
    End If
End Function

Public Function MontaCabLV(Cab0 As String, Cab1 As String, Cab2 As String, Cab3 As String, Cab4 As String, Cab5 As String, Cab6 As String, Cab7 As String, Cab8 As String, Cab9 As String, Cab10 As String, Cab11 As String, Cab12 As String, Cab13 As String, Cab14 As String, Cab15 As String, Cab16 As String, Cab17 As String, Cab18 As String, Cab19 As String, Cab20 As String, Cab21 As String, Cab22 As String, Cab23 As String, Cab24 As String, Cab25 As String)
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
    NomeColLV(21) = Cab21
    NomeColLV(22) = Cab22
    NomeColLV(23) = Cab23
    NomeColLV(24) = Cab24
    NomeColLV(25) = Cab25
End Function

Public Function DimensionaLV(NomeLV As String)
    MeuLV.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
    MeuLV.Frame1.Caption = NomeLV
    MeuLV.Frame1.Move 0, 0, Principal.ScaleWidth - 300, Principal.ScaleHeight - 650
    MeuLV.ListView1.Move 100, 1000, Principal.ScaleWidth - 500, Principal.ScaleHeight - 1800
End Function

Public Function DimensionaForm()
'    Dim XSize As Integer
'    frmMonitorar.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
'    frmMonitorar.Frame2.Move 4680, 120, Principal.ScaleWidth - 5000, Principal.ScaleHeight - 800
'    frmMonitorar.ListView1.Move 120, 240, Principal.ScaleWidth - 11300, Principal.ScaleHeight - 1200
'    If FimAprop = "N" Then
'        frmMonitorar.ListView3.Move 120 + Principal.ScaleWidth - 11200, 240, frmMonitorar.Frame12.Width + 700, Principal.ScaleHeight - 1200
'    Else
'        frmMonitorar.ListView3.Move 120 + Principal.ScaleWidth - 11200, 240, frmMonitorar.Frame12.Width + 700, Principal.ScaleHeight - 2000
'    End If
'    frmMonitorar.Command1.Move Principal.ScaleWidth - 2800, Principal.ScaleHeight - 1500
End Function

Public Function DimensionaPPS()
'    frmProgramacao.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
'    frmProgramacao.ListView1.Move 120, frmProgramacao.Top + 10, Principal.ScaleWidth - 500, Principal.ScaleHeight - 1800
'    frmProgramacao.Frame1.Move 7020, frmProgramacao.Height - 1700, 3975, 1095
'    frmProgramacao.Frame2.Move 11100, frmProgramacao.Height - 1700, 6615, 1095
'    frmProgramacao.Frame3.Move 4400, frmProgramacao.Height - 1700, 2535, 1095
'    frmProgramacao.cmdCadastro(12).Move 120, frmProgramacao.Height - 1250, 615, 615
'    frmProgramacao.cmdCadastro(13).Move 735, frmProgramacao.Height - 1250, 615, 615
'    frmProgramacao.cmdCadastro(0).Move 1350, frmProgramacao.Height - 1250, 615, 615
'    frmProgramacao.Frame4.Move 2120, frmProgramacao.Height - 1700, 1935, 1095
End Function

Public Function MontaCabecalhoLV()
    Dim X As Integer
    'Limpa o cabeçalho antes de compor novamente
    MeuLV.ListView1.ColumnHeaders.Clear
    With MeuLV.ListView1.ColumnHeaders
        For X = 0 To 25
            If NomeColLV(X) = "" Then Exit Function
            .Add , , NomeColLV(X), Len(NomeColLV(X)) * 144
            QtdColReal = QtdColReal + 1
        Next
    End With
End Function

Public Function MontaDadosLV(ZeroPriCol As String)
'On Error GoTo Err
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
    Exit Function
End Function

Public Function PersonaColLV(posCol As Integer, negritoCol As String, corCol As String, caracterCol As String, imageCol As String, formataColZero As String, formataColDecimal As String, alinhaCol As String)
    On Error Resume Next
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        Principal.ProgressBar1.Value = X
        Set ItemLst = MeuLV.ListView1.ListItems.Item(X)
        'NEGRITO NOS ITENS DA COLUNA
        If negritoCol = "S" Then ItemLst.ListSubItems(posCol).Bold = True
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
                    ItemLst.ListSubItems(7).ForeColor = &HFF&
                    ItemLst.ListSubItems(8).ForeColor = &HFF&
                    ItemLst.ListSubItems(9).ForeColor = &HFF&
                    ItemLst.ListSubItems(10).ForeColor = &HFF&
                    ItemLst.ListSubItems(11).ForeColor = &HFF&
                End If
                If ItemLst.ListSubItems(7) >= MediaGlobal Then
                    ItemLst.ListSubItems(7).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(7) < MediaGlobal And ItemLst.ListSubItems(7) >= vAprovadoRest Then
                    ItemLst.ListSubItems(7).ForeColor = &H80FF&
                Else
                    ItemLst.ListSubItems(7).ForeColor = &HC0&
                End If
            ElseIf Formulario = "Empréstimo/Devolução" Then
                If ItemLst.ListSubItems(posCol) = "Sim" Then 'LARANJA
                    ItemLst.ForeColor = &H80FF&
                    ItemLst.ListSubItems(1).ForeColor = &H80FF&
                    ItemLst.ListSubItems(2).ForeColor = &H80FF&
                    ItemLst.ListSubItems(3).ForeColor = &H80FF&
                    ItemLst.ListSubItems(4).ForeColor = &H80FF&
                    ItemLst.ListSubItems(5).ForeColor = &H80FF&
                    ItemLst.ListSubItems(6).ForeColor = &H80FF&
                End If
            ElseIf Formulario = "Fornecedores" Then
                If ItemLst.ListSubItems(posCol) = "Credenciado" And ItemLst.ListSubItems(posCol + 2) <> "-" Then 'VERDE
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
                    ItemLst.ListSubItems(13).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(posCol) = "-" Then 'VERMELHO
                    ItemLst.ForeColor = &HC0&
                    ItemLst.ListSubItems(1).ForeColor = &HC0&
                    ItemLst.ListSubItems(2).ForeColor = &HC0&
                    ItemLst.ListSubItems(3).ForeColor = &HC0&
                    ItemLst.ListSubItems(4).ForeColor = &HC0&
                    ItemLst.ListSubItems(5).ForeColor = &HC0&
                    ItemLst.ListSubItems(6).ForeColor = &HC0&
                    ItemLst.ListSubItems(7).ForeColor = &HC0&
                    ItemLst.ListSubItems(8).ForeColor = &HC0&
                    ItemLst.ListSubItems(9).ForeColor = &HC0&
                    ItemLst.ListSubItems(10).ForeColor = &HC0&
                    ItemLst.ListSubItems(11).ForeColor = &HC0&
                    ItemLst.ListSubItems(12).ForeColor = &HC0&
                    ItemLst.ListSubItems(13).ForeColor = &HC0&
                ElseIf ItemLst.ListSubItems(posCol) = "Credenciado" And ItemLst.ListSubItems(posCol + 2) = "-" Then 'LARANJA
                    ItemLst.ForeColor = &H80FF&
                    ItemLst.ListSubItems(1).ForeColor = &H80FF&
                    ItemLst.ListSubItems(2).ForeColor = &H80FF&
                    ItemLst.ListSubItems(3).ForeColor = &H80FF&
                    ItemLst.ListSubItems(4).ForeColor = &H80FF&
                    ItemLst.ListSubItems(5).ForeColor = &H80FF&
                    ItemLst.ListSubItems(6).ForeColor = &H80FF&
                    ItemLst.ListSubItems(7).ForeColor = &H80FF&
                    ItemLst.ListSubItems(8).ForeColor = &H80FF&
                    ItemLst.ListSubItems(9).ForeColor = &H80FF&
                    ItemLst.ListSubItems(10).ForeColor = &H80FF&
                    ItemLst.ListSubItems(11).ForeColor = &H80FF&
                    ItemLst.ListSubItems(12).ForeColor = &H80FF&
                    ItemLst.ListSubItems(13).ForeColor = &H80FF&
                End If
            ElseIf Formulario = "Recebimento de Ordem de Compra" Then
                If ItemLst.ListSubItems(posCol) = "S" Then 'Verde
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
                    ItemLst.ListSubItems(13).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(posCol) = "N" Then 'Cinza
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
                    ItemLst.ListSubItems(13).ForeColor = &H808080
                ElseIf ItemLst.ListSubItems(posCol) = "7" Then 'Laranja
                    ItemLst.ForeColor = &H80FF&
                    ItemLst.ListSubItems(1).ForeColor = &H80FF&
                    ItemLst.ListSubItems(2).ForeColor = &H80FF&
                    ItemLst.ListSubItems(3).ForeColor = &H80FF&
                    ItemLst.ListSubItems(4).ForeColor = &H80FF&
                    ItemLst.ListSubItems(5).ForeColor = &H80FF&
                    ItemLst.ListSubItems(6).ForeColor = &H80FF&
                    ItemLst.ListSubItems(7).ForeColor = &H80FF&
                    ItemLst.ListSubItems(8).ForeColor = &H80FF&
                    ItemLst.ListSubItems(9).ForeColor = &H80FF&
                    ItemLst.ListSubItems(10).ForeColor = &H80FF&
                    ItemLst.ListSubItems(11).ForeColor = &H80FF&
                    ItemLst.ListSubItems(12).ForeColor = &H80FF&
                    ItemLst.ListSubItems(13).ForeColor = &H80FF&
                End If
                
                If ItemLst.ListSubItems(posCol) = "A" Then 'VERDE
                    ItemLst.ListSubItems(6).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(posCol) = "B" Then 'LARANJA
                    ItemLst.ListSubItems(6).ForeColor = &H80FF&
                ElseIf ItemLst.ListSubItems(posCol) = "C" Then 'VERMELHO
                    ItemLst.ListSubItems(6).ForeColor = &HC0&
                End If
            Else
                If ItemLst.ListSubItems(posCol) >= MediaGlobal Then
                    ItemLst.ListSubItems(posCol).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(posCol) < MediaGlobal And ItemLst.ListSubItems(posCol) >= vAprovadoRest Then
                    ItemLst.ListSubItems(posCol).ForeColor = &H80FF&
                Else
                    ItemLst.ListSubItems(posCol).ForeColor = &HC0&
                End If
            End If
        End If
        'CASAS DECIMAIS NOS ITENS DA COLUNA
        If formataColDecimal = "S" Then ItemLst.SubItems(posCol) = "" & Format(ItemLst.SubItems(posCol), "#,##0.00;(#,##0.00)")
        'FORMATAÇÃO DE 6 ZEROS NOS ITENS DA COLUNA
        If formataColZero = "S" Then ItemLst.SubItems(posCol) = "" & Format(ItemLst.SubItems(posCol), "000000")
        'ADICIONAR CARACTER(ES) NOS ITENS DA COLUNA
        If caracterCol <> "" Then ItemLst.SubItems(posCol) = ItemLst.ListSubItems(posCol) & caracterCol
        'INFORMA SE IRÁ UTILIZAR O IMAGELIST NOS ITENS DA COLUNA
        If imageCol = "S" Then
            'A condição abaixo verifica o conteudo da posição do Listview
            If ItemLst.SubItems(posCol) = "S" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
            ElseIf ItemLst.SubItems(posCol) = "1" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
                If apontaLV = 0 Then
                    ItemLst.ListSubItems(10) = "EXPORTADO"
                    ItemLst.ListSubItems(10).ForeColor = &HC0&
                Else
                    ItemLst.ListSubItems(8) = "EXPORTADO"
                    ItemLst.ListSubItems(8).ForeColor = &HC0&
                End If
            ElseIf ItemLst.SubItems(posCol) = "2" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "EXC"
                If apontaLV = 0 Then
                    ItemLst.ListSubItems(10) = "CANCELADO"
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
                    ItemLst.ListSubItems(13).ForeColor = &H808080
                    ItemLst.ListSubItems(14).ForeColor = &H808080
                    ItemLst.ListSubItems(15).ForeColor = &H808080
                    ItemLst.ListSubItems(16).ForeColor = &H808080
                Else
                    ItemLst.ListSubItems(8) = "CANCELADO"
                    ItemLst.ListSubItems(8).ForeColor = &H808080
                End If
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
            ElseIf ItemLst.SubItems(posCol) = "E" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "EMPRESTAR"
            ElseIf ItemLst.SubItems(posCol) = "D" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "DEVOLVER"
            Else
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "AST"
            End If
            ItemLst.SubItems(posCol) = ""
            
            'CONDIÇÃO ESPECIAL PARA O SISTEMA DE MEDIÇÃO
            If apontaLV = 0 Then
                If ItemLst.ListSubItems(10) = "Reprovado" Then
                    ItemLst.ListSubItems.Item(14).ReportIcon = "EXC1"
                ElseIf ItemLst.ListSubItems(10) = "Aguardando Aprovação" Then
                    ItemLst.ListSubItems.Item(14).ReportIcon = "APR"
                ElseIf ItemLst.ListSubItems(10) = "Aprovação Parcial" Then
                    ItemLst.ListSubItems.Item(14).ReportIcon = "APP"
                End If
            Else
                If ItemLst.ListSubItems(8) = "Reprovado" Then
                    ItemLst.ListSubItems.Item(16).ReportIcon = "EXC1"
                ElseIf ItemLst.ListSubItems(8) = "Aguardando Aprovação" Then
                    ItemLst.ListSubItems.Item(16).ReportIcon = "APR"
                ElseIf ItemLst.ListSubItems(8) = "Aprovação Parcial" Then
                    ItemLst.ListSubItems.Item(16).ReportIcon = "APP"
                End If
            End If
            
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
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Function

Public Function PersonaColLVForm(LVForm As ListView, posCol As Integer, negritoCol As String, corCol As String, caracterCol As String, imageCol As String, formataColZero As String, formataColDecimal As String, alinhaCol As String)
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = LVForm.ListItems.Count
    For X = 1 To Y
        Set ItemLst = LVForm.ListItems.Item(X)
        'NEGRITO NOS ITENS DA COLUNA
        If negritoCol = "S" Then ItemLst.ListSubItems(posCol).Bold = True
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
                End If
                If ItemLst.ListSubItems(7) >= MediaGlobal Then
                    ItemLst.ListSubItems(7).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(7) < MediaGlobal And ItemLst.ListSubItems(7) >= vAprovadoRest Then
                    ItemLst.ListSubItems(7).ForeColor = &H80FF&
                Else
                    ItemLst.ListSubItems(7).ForeColor = &HC0&
                End If
            ElseIf Formulario = "Empréstimo/Devolução" Then
                If ItemLst.ListSubItems(posCol) = "Sim" Then 'VERMELHO
                    On Error Resume Next
                    ItemLst.ForeColor = &HC0&
                    ItemLst.ListSubItems(1).ForeColor = &HC0&
                    ItemLst.ListSubItems(2).ForeColor = &HC0&
                    ItemLst.ListSubItems(3).ForeColor = &HC0&
                    ItemLst.ListSubItems(4).ForeColor = &HC0&
                    ItemLst.ListSubItems(5).ForeColor = &HC0&
                    ItemLst.ListSubItems(6).ForeColor = &HC0&
                    ItemLst.ListSubItems(7).ForeColor = &HC0&
                    ItemLst.ListSubItems(8).ForeColor = &HC0&
                    ItemLst.ListSubItems(9).ForeColor = &HC0&
                    ItemLst.ListSubItems(10).ForeColor = &HC0&
                End If
            ElseIf Formulario = "Método & Processo" Then
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
                End If
            Else
                If ItemLst.ListSubItems(posCol) >= MediaGlobal Then
                    ItemLst.ListSubItems(posCol).ForeColor = &H8000&
                ElseIf ItemLst.ListSubItems(posCol) < MediaGlobal And ItemLst.ListSubItems(posCol) >= vAprovadoRest Then
                    ItemLst.ListSubItems(posCol).ForeColor = &H80FF&
                Else
                    ItemLst.ListSubItems(posCol).ForeColor = &HC0&
                End If
            End If
        End If
        'CASAS DECIMAIS NOS ITENS DA COLUNA
        If formataColDecimal = "S" Then ItemLst.SubItems(posCol) = "" & Format(ItemLst.SubItems(posCol), "#,##0.00;(#,##0.00)")
        'FORMATAÇÃO DE 6 ZEROS NOS ITENS DA COLUNA
        If formataColZero = "S" Then ItemLst.SubItems(posCol) = "" & Format(ItemLst.SubItems(posCol), "000000")
        'ADICIONAR CARACTER(ES) NOS ITENS DA COLUNA
        If caracterCol <> "" Then ItemLst.SubItems(posCol) = ItemLst.ListSubItems(posCol) & caracterCol
        'INFORMA SE IRÁ UTILIZAR O IMAGELIST NOS ITENS DA COLUNA
        If imageCol = "S" Then
            'A condição abaixo verifica o conteudo da posição do Listview
            If ItemLst.SubItems(posCol) = "S" And ItemLst.ListSubItems.Item(posCol).ReportIcon <> "OK" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
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
            ElseIf ItemLst.SubItems(posCol) = "N" And ItemLst.ListSubItems.Item(posCol).ReportIcon <> "EXC" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "EXC"
            ElseIf ItemLst.SubItems(posCol) = "E" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "EMPRESTADO"
            ElseIf ItemLst.SubItems(posCol) = "D" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "DEVOLVIDO"
            End If
            ItemLst.SubItems(posCol) = ""
        End If
        'ALINHAMENTO DA COLUNA
        If alinhaCol = "D" Then
            LVForm.ColumnHeaders(posCol + 1).Alignment = lvwColumnRight
        ElseIf alinhaCol = "E" Then
            LVForm.ColumnHeaders(posCol + 1).Alignment = lvwColumnLeft
        Else
            LVForm.ColumnHeaders(posCol + 1).Alignment = lvwColumnCenter
        End If
    Next
End Function


Public Sub ExcluirDadosLV(QualLV As Integer)
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim rsExcLVGeral As New ADODB.Recordset
    cnBanco.BeginTrans
    mobjMsg.Abrir "Confirma exclusão da " & LegendaExc & " selecionada?", YesNo, pergunta, "IMRM"
    If Tp = 1 Then
        'Módulo de Medições Teceiros
        If QualLV = 0 Then
            Msgbox "Em desenvolvimento"
            'Verifica se existe alguma registro na tabela tbCriterioSubRec dependente do registro da tabela tbCriterioRec
'            SqlExcLVGeral = "select * from tbGrupoCriterioItens where idcriteriorec = '" & Val(varGlobal) & "'"
'            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
'            If rsExcLVGeral.RecordCount > 0 Then
'                'Não irá excluir, irá desativar
'                rsExcLVGeral.Close
'                Set rsExcLVGeral = Nothing
'                SqlExcLVGeral = "Update tbCriterioRec set ativo = 'N' where idcriteriorec = '" & Val(varGlobal) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "Registro desativado com sucesso", Ok, informacao, "IMRM"
'            Else
'                'Irá excluir
'                rsExcLVGeral.Close
'                Set rsExcLVGeral = Nothing
'                SqlExcLVGeral = "Delete from tbCriterioRec where idcriteriorec= '" & Val(varGlobal) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "Registro excluido com sucesso", Ok, informacao, "IMRM"
'            End If
        'Módulo Medições PJ
        ElseIf QualLV = 1 Then
            Msgbox "Em desenvolvimento"
'            'Verifica se existe alguma registro na tabela tbFornecedores dependente do registro da tabela tbAvFornecGrup
'            SqlExcLVGeral = "select * from tbFornecedores as a inner join tbAvFornecGrup as b on a.grupo = b.nomeavfornecgrup where idavfornecgrup = '" & Val(varGlobal) & "'"
'            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
'            If rsExcLVGeral.RecordCount > 0 Then
'                'Não irá excluir, irá desativar
'                rsExcLVGeral.Close
'                Set rsExcLVGeral = Nothing
'                SqlExcLVGeral = "Update tbAvFornecGrup set ativo = 'N' where idavfornecgrup = '" & Val(varGlobal) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "Registro desativado com sucesso", Ok, informacao, "IMRM"
'            Else
'                'Irá excluir
'                rsExcLVGeral.Close
'                Set rsExcLVGeral = Nothing
'                SqlExcLVGeral = "Delete from tbAvFornecGrup where idavfornecgrup= '" & Val(varGlobal) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                SqlExcLVGeral = "Delete from tbAvFornecGrupItens where idavfornecgrup= '" & Val(varGlobal) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "Grupo excluido com sucesso", Ok, informacao, "IMRM"
'            End If
        'Módulo Usuários
        ElseIf QualLV = 13 Then
            'Verifica se existe o critério na tabela tbAvFornecGrupItens
            'SqlExcLVGeral = "select * from tbAvFornecGrupItens where idcriterioavfornec = '" & Val(varGlobal) & "'"
            'rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
            'If rsExcLVGeral.RecordCount > 0 Then
            '    'Não irá excluir, irá desativar
            '    rsExcLVGeral.Close
            '    Set rsExcLVGeral = Nothing
            '    SqlExcLVGeral = "Update tbCriterioAvFornec set ativo = 'N' where idcriterioavfornec = '" & Val(varGlobal) & "'"
            '    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
            '    mobjMsg.Abrir "Registro desativado com sucesso", Ok, informacao, "IMRM"
            'Else
                'Irá excluir
                'rsExcLVGeral.Close
                'Set rsExcLVGeral = Nothing
                SqlExcLVGeral = "Delete from tbusuarios where codigo= '" & Val(varGlobal) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                SqlExcLVGeral = "Delete from tbsenha where codigo= '" & Val(varGlobal) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                SqlExcLVGeral = "Delete from TBLOCALESTOQUE where codigo= '" & Val(varGlobal) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                
                mobjMsg.Abrir "usuário excluido com sucesso", Ok, informacao, "IMRM"
            'End If
        End If
    End If
    cnBanco.CommitTrans
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Public Sub ExcluirItemLV(LV As ListView)
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
Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
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
                    SendMessage LV.hwnd, LVM_FIRST + 30, c.Index - 1, -1
                    
                    If Mid$(MeuLV.ListView1.ColumnHeaders.Item(posi).Width, 1, 3) = 180 Then
                        MeuLV.ListView1.ColumnHeaders.Item(posi).Width = 0
                    End If
                End If
                posi = posi + 1
            End If
        Next
    Else
        SendMessage LV.hwnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
    Exit Sub
Err:
    Resume Next
End Sub

'ROTINAS/FUNÇÕES DO LISTVIEW GENERICO - DAKI PARA CIMA
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Avaliador(Tipo As String)
'On Error GoTo Err
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
        If PontosColabExp >= MediaGlobal Then
            chamaForm.Label37.ForeColor = &H8000&
        ElseIf PontosColabExp < MediaGlobal And PontosColabExp >= vAprovadoRest Then
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
        If PontosTotaisHab >= MediaGlobal Then
            chamaForm.Label38.ForeColor = &H8000&
        ElseIf PontosTotaisHab < MediaGlobal And PontosTotaisHab >= vAprovadoRest Then
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
            If PontosTotaisTrei >= MediaGlobal Then
                chamaForm.Label39.ForeColor = &H8000&
            ElseIf PontosTotaisTrei < MediaGlobal And PontosTotaisTrei >= vAprovadoRest Then
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
        If PontosColabFor >= MediaGlobal Then
            chamaForm.Label40.ForeColor = &H8000&
        ElseIf PontosColabFor < MediaGlobal And PontosColabFor >= vAprovadoRest Then
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
        
            If PontosColabADP >= MediaGlobal Then
                chamaForm.Label43.ForeColor = &H8000&
            ElseIf PontosColabADP < MediaGlobal And PontosColabADP >= vAprovadoRest Then
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
        If ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor + PontosColabADP) / Contador) >= MediaGlobal Then
            chamaForm.Label41.ForeColor = &H8000&
        ElseIf ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor) + PontosColabADP / Contador) < MediaGlobal And ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor) + PontosColabADP / Contador) >= vAprovadoRest Then
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
    Exit Function
End Function

Public Function gravaLog(Campo1 As String, campo2 As String, campo3 As String)
    If GeraLog = "N" Then Exit Function
    Dim sqlLog As String
    Dim rsLog As New ADODB.Recordset
    
    sqlLog = "Insert into tbLog(data,hora,usuario,grupo,formulario,acao,codcoligada) Values('" & CStr(Date) & "','" & CStr(Time) & "','" & NomUsu & "','" & GrupoUsu & "','" & Formulario & "','" & Pesquisa & ":" & Campo1 & "-" & campo2 & "-" & campo3 & "','" & vCodcoligada & "')"
    rsLog.Open sqlLog, cnBanco

End Function

Public Function gravaSolicitacao(vCPF As String, vTipo As String, vNota As String, vSolicitacao As String, vSolicitante As String)
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

End Function

Public Function caculaTmpExp()
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
End Function

Public Sub ajusta_LV()
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
End Sub

Public Sub GravarConfLV()
    'ESSA ROTINA RESTAURA AS CONFIGURAÇÕES DE POSICIONAMENTO E TAMANHO DAS COLUNAS
    'DEFINIDAS PELO USUÁRIO.
    'A TABELA TBCONFLV ARMAZENA AS CONFIGURAÇÕES DE POSICIONAMENTO E TAMANHO DAS COLUNAS.
    
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim X As Integer, Y As Integer
    'cnBanco.BeginTrans
   
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
    'cnBanco.CommitTrans
    rsSalvar.Close
    Set rsSalvar = Nothing
End Sub

Public Function Enter(Key As Integer) As Integer
   If Key = 13 Then
       Enter = 0
   Else
       Enter = Key
   End If
End Function

Public Function NovoCodigo()
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
End Function

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'A Função abaixo gera código para qualquer Listview
Public Function GeraCodigoLV(LV As ListView)
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
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    If vCampo2 = "" Then
        SqlGera = "Select top 1 * from " & vTabela & " order by " & vCampo & " Desc"
    Else
        SqlGera = "Select top 1 * from " & vTabela & " where " & vCampo2 & "=" & Val(vText) & " order by " & vCampo & " Desc"
    End If
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraCodigo.RecordCount > 0 Then
        GeraCodigoTB = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigoTB = 1
    End If
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
End Function

'A Função abaixo chama grid para quaisquer: textbox e tabela
Public Function ChamaGrid(vTabela As String, vCampo As String, vTxt As TextBox, vForm As Form, vPesq1 As String, vPesq2 As String)
    Dim F As New frmpesqger
    Dim Iposicao As Variant
'    If vTabela = "tbGrupoClass" Then
'        Sqlp = "Select " & vPesq1 & "," & vPesq2 & " from " & vTabela & " where idprd ='" & frmFormulaCC.txtformula(0) & "' order by " & vCampo & ""
'    ElseIf vTabela = "CORPORERM.dbo.GCCUSTO" Then
'        Sqlp = "select a.CODREDUZIDO,a.NOME from CORPORERM.dbo.GCCUSTO as a left join ZEUS.dbo.tbFormula as b on a.CODREDUZIDO = b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS " & _
'            "Where a.ATIVO  = 'T' and b.nmform is not null group by a.ID,a.CODREDUZIDO,a.NOME order by a.CODREDUZIDO"
'    Else
        Sqlp = "Select " & vPesq1 & "," & vPesq2 & " from " & vTabela & "  order by " & vCampo & ""
'    End If
    procnom = vCampo
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa"
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
End Function

'A Função abaixo é referente a ALTERAÇÃO de dados de qualquer ListView com até 15 colunas
Public Function AlteraLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
    Dim Y As Integer, X As Integer, Z As Integer
    Dim vRaptor(15) As String
    For X = LBound(vRaptor) To UBound(vRaptor)
        vRaptor(X) = ""
    Next X
    Y = LV.ListItems.Count
    If Y = 0 Then Exit Function
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    
    'SOMENTE PARA ZEUS--
    If apontaLV = 9 Then
        '1º VERIFICA SE A DATA PREVISTA ESTA VAZIA
        If LV.SelectedItem.ListSubItems(5).Text <> "" And LV.SelectedItem.ListSubItems(5).Text <> "-" Then
            '2º VERIFICA SE A SEMANA ATUAL É MAIOR OU IGUAL A SEMANA PROGRAMADA
            If DatePart("ww", (Date)) >= DatePart("ww", CDate(LV.SelectedItem.ListSubItems(5).Text)) Then
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
    '-------------------
    
    For Z = 1 To LV.ColumnHeaders.Count
        If Z = 1 Then
            vRaptor(Z) = LV.ListItems.Item(X)
        Else
            vRaptor(Z) = LV.SelectedItem.ListSubItems.Item(Z - 1)
        End If
    Next
    If vRaptor(1) <> "" Then vCP01.Text = vRaptor(1)
    If vRaptor(2) <> "" Then vCP02.Text = vRaptor(2)
    If vRaptor(3) <> "" Then vCP03.Text = vRaptor(3)
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
End Function

Private Sub bloqueiaEdicaoMP(vTipo As Boolean)
    'Dim X As Integer
    
    'chamaForm.TreeView1.Enabled = vTipo
    'chamaForm.TreeView2.Enabled = vTipo
    'chamaForm.TreeView3.Enabled = vTipo
    'For X = 0 To chamaForm.cmdCadastro.Count - 1
    '    chamaForm.cmdCadastro(X).Enabled = vTipo
    'Next
    'chamaForm.cmdCadastro(13).Enabled = vTipo
    'chamaForm.txtformula(0).Enabled = vTipo
    'chamaForm.txtformula(5).Enabled = vTipo
    'chamaForm.txtformula(12).Enabled = vTipo
    'chamaForm.txtformula(13).Enabled = vTipo
    'chamaForm.txtformula(26).Enabled = vTipo
    'chamaForm.Combo1.Enabled = vTipo
    'chamaForm.SSTab1.TabEnabled(2) = vTipo
    'chamaForm.DTPicker1.Enabled = vTipo
    'chamaForm.DTPicker2.Enabled = vTipo
    'chamaForm.SkinLabel20.Visible = vTipo
End Sub

'A Função abaixo é referente a INCLUSÃO de dados de qualquer ListView com até 10 colunas
Public Function IncluirLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
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
'                Exit Function
            End If
        Next
        Set ItemLst = LV.ListItems.Add(, , vRaptor(1))
        Y = LV.ListItems.Count
    Else
'        If chamaForm.Name = "frmMPCompleto" And LV.Name = "ListView1" Then
'            If separaDesLv(chamaForm.Text1.Text) = False Then
'                IncluirLV = False
'                Exit Function
'            Else
'                IncluirLV = True
'            End If
'        End If
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
'            vCodLM = Mid$(RECEBE, 1, 2)
'            vCodSeq = Mid$(RECEBE, 3, 3)
            SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
            "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoSAP & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
            "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoSAP & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
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
'        vCodLM = Mid$(RECEBE, 1, 2)
'        vCodSeq = Mid$(RECEBE, 3, 3)
        SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
        "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoSAP & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
        "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoSAP & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
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
End Function

'A Função abaixo é referente a VALIDAÇÃO de dados de qualquer ListView com até 10 colunas
Public Function ValidaCampos(LV As ListView, vTxt1 As TextBox, vTxt2 As TextBox, vTxt3 As TextBox, vTxt4 As TextBox, vTxt5 As TextBox, vTxt6 As TextBox, vTxt7 As TextBox, vTxt8 As TextBox, vTxt9 As TextBox, vTxt10 As TextBox, vTxt11 As TextBox, vTxt12 As TextBox, vTxt13 As TextBox, vTxt14 As TextBox, vTxt15 As TextBox)
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
Public Function LimpaLV(LV As ListView)
    LV.ListItems.Clear
End Function

'A Função abaixo preenche dados da TXT2 baseado no dado informado na TXT1
Public Function CarregaTxt(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vCampo2 As String, vTipoCampo2 As String, vVar1 As TextBox, vVar2 As TextBox, vPosicao1 As Integer, vPosicao2 As Integer, vRetorno1 As TextBox, vTipoRetorno1 As String, vRetorno2 As TextBox, vQualQuery As String)
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
        mobjMsg.Abrir "Dado não cadastrado", Ok, informacao, "Atenção"
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
End Function

'A Função abaixo grava dados de até 50 variáveis de um formulário em uma determinada tabela
Public Function GravaDados(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vVar1 As TextBox, vQtdCampos As Integer, vCampo2 As String, vTipoCampo2 As String, vVar2 As TextBox)
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
    If vTransacaoAtiva = 0 Then cnBanco.BeginTrans
   
    If vTipoCampo1 = "I" Then
        If vCampo2 = "" Then
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = " & Val(vVar1) & ""
        Else
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "' and " & vCampo2 & " = '" & Val(vVar2) & "'"
        End If
    Else
        If vCampo2 = "" Then
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "'"
        Else
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "' and " & vCampo2 & " = '" & vVar2 & "'"
        End If
    End If
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvar.EOF Then rsSalvar.AddNew
    For X = 0 To vQtdCampos - 1
        If vQualquerDado(X + 1, 2) = "S" Then
            rsSalvar.Fields(X) = vQualquerDado(X + 1, 1)
        ElseIf vQualquerDado(X + 1, 2) = "I" Then
            rsSalvar.Fields(X) = Val(vQualquerDado(X + 1, 1))
        ElseIf vQualquerDado(X + 1, 2) = "D" Then
            rsSalvar.Fields(X) = CDate(vQualquerDado(X + 1, 1))
        End If
    Next
    rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    If vTransacaoAtiva = 0 Then cnBanco.CommitTrans
    'MsgBox "Os dados do FORMULARIO foram salvos com sucesso", vbInformation, "ProtótipoX"
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

Public Function ordenaLVArray(LV As ListView, vPos0 As String, vPos1 As String, vPos2 As String, vPos3 As String, vPos4 As String, vPos5 As String, vPos6 As String, vPos7 As String, vPos8 As String, vPos9 As String, vPos10 As String, vPos11 As String, vPos12 As String, vPos13 As String, vPos14 As String, vPos15 As String)
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
On Error Resume Next
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer, X As Integer
    
    cnBanco.BeginTrans
    
    If vCampo1 = "" Then
        If apontaLV = 51 Or apontaLV = 1 Then
            sqlDeletar = Sqlp
        Else
            sqlDeletar = "Delete from " & vTabela
        End If
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
        While vQualquerDado(X, Y) <> "" 'Or vQualquerDado(X, Y) <> " "
            If apontaLV = 9 And Y = 10 Then 'Essa condição serve Somente para o zeus
                rsSalvar.Fields(9) = Mid$(vQualquerDado(X, Y), 1, 9)
                rsSalvar.Fields(15) = Mid$(vQualquerDado(X, Y), 11, 1)
            Else
                If vQualquerDado(X, Y) <> "-" Then rsSalvar.Fields(Y - 1) = vQualquerDado(X, Y)
            End If
            Y = Y + 1
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
TrataErro:
    Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbCritical, "Atenção"
    cnBanco.RollbackTrans
    Exit Function
End Function

Public Function chamaSQL(vSql As String)
    Sqlp = ""
    Sqlp = vSql
End Function

Public Function Compoe_Listview(LV As ListView, vSqlCompoe As String, vZerosEsq As String)
    ' Declaração de variaveis
    Dim rsCompoe As New ADODB.Recordset
    rsCompoe.Open vSqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly

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
End Function

Public Function Compoe_ListviewFiltro(LV As ListView, vSqlCompoe As String, vZerosEsq As String)
    ' Declaração de variaveis
    Dim rsCompoe As New ADODB.Recordset
    rsCompoe.Open vSqlCompoe, cnBancoSAP, adOpenKeyset, adLockReadOnly

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
End Function

Public Function mudaCorText(txt As TextBox)
    txt.BackColor = &HC0E0FF
End Function

Public Function voltaCorText(txt As TextBox)
    txt.BackColor = &HFFFFFF
End Function

Public Function converteSemana(vSemanaAno As Integer, vDataAno As DTPicker, vAnoDaSemana As String)
    On Error GoTo Err
    Dim vDataConvertida As String 'Dia/Mes/Ano
    Dim vAno As String, vMes As String, vDia As String
    Dim vX As Integer, vY As Integer
    
    If vAnoDaSemana = "" Then
        vAno = DatePart("yyyy", Date)
    Else
        vAno = vAnoDaSemana
    End If
    
    For vX = 1 To 12
        vMes = Format(vX, "00")
        For vY = 1 To 31
            vDia = Format(vY, "00")
            vDataConvertida = Format(vDia & "/" & vMes & "/" & vAno, "dd/mm/yyyy")
            If IsDate(CDate(vDataConvertida)) Then
                If DatePart("ww", CDate(vDataConvertida)) = vSemanaAno Then
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

Public Function MarcaDesmarcaTodos(LV As ListView)
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

Public Function MarcaDesmarcaGeral(LV As ListView)
    'Deixa checado somente um item do Listview
    Dim X As Integer, Y As Integer, J As Integer, vPosicaoAtual As Integer
    Y = LV.ListItems.Count
    If Y = 0 Then Exit Function
    J = LV.SelectedItem.Index
    For X = 1 To Y
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        End If
    Next
    LV.ListItems.Item(J).Checked = True
    vPosicaoAtual = J
End Function

Public Function SomaLV(LV As ListView, vColunaLV As Integer, vTxtRetorno As TextBox)
    On Error Resume Next
    Dim X As Integer, Y As Integer, F As Integer
    Y = LV.ListItems.Count
    If Y = 0 Then
        vTxtRetorno = 0
    End If
    Dim somaTempo As Double
    somaTempo = 0
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then F = X
    Next
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        'If Trim$(LV.SelectedItem.ListSubItems.Item(6)) <> " " Then
            somaTempo = somaTempo + LV.SelectedItem.ListSubItems.Item(vColunaLV)
        'End If
    Next
    If somaTempo <> 0 Then
        vTxtRetorno.Text = Format(somaTempo, "#,##00.00;(#,##0.00)")
        LV.ListItems.Item(F).Selected = True
    End If
End Function

Public Function MarcaDesmarca(LV As ListView)
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
