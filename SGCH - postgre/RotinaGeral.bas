Attribute VB_Name = "RotinaGeral"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Biblioteca para manipulação do Regedit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000

Public oConn As ADODB.Connection
Public sDatabaseName As String 'Utilizado para criar nova conexão com o banco na tela de splash
Public sServerName As String ' Utilizado para criar nova conexão com o banco na tela de splash
Public sUsuName As String 'Nome do usuário de conexão ao DB
Public sSenhaDB As String 'Senha de conexão ao DB
Public sSGBD As Integer 'Versão do SGBD
Public vFormatoDatetime As String

Public sLogoEmpresa As String ' Utilizado para guardar o caminho da logo da empresa
Public StatusTrei As String 'Verifica o status do treinamento

Public rsVExp As New ADODB.Recordset
Public SqlVExp As String

Public cnBanco As ADODB.Connection
Public cnBancoTotvs As ADODB.Connection
'Public rsResumo As New ADODB.Recordset
Public MediaGlobal As Double
Public vAprovadoRest As Double
Public vAvisos As String 'Ao entrar o sistema é exibida uma tela de Avisos, onde será informado pendências no sistema
Public vCalcExp As String 'Calcula automaticamente o tempo de experiência dos colaboradores
Public GeraIntr As String 'Identifica se o sistema irá gerar ou não treinamentos introdutorios para colaboradores
Public GeraObri As String 'Identifica se o sistema irá gerar ou não treinamentos obrigatorios para colaboradores
Public GeraLog As String

Public vAfastDias As String
Public vAfastTreiInt As String
Public vAfastTreiObr As String

Public XCodGrp As Integer 'Armazena o codigo do grupo que o usuário esta logado
Public vInc As String, vExc As String, vEdi As String, vSal As String, vImp As String, vFil As String, vAva As String, vAdi As String, vDem As String, vAdiRep As String, vAdiRes As String

Public varGlobal As String
Public varGlobal2 As String
Public FiltroGeral As String
Public Formulario As Variant
Public Sqlp As String
Public SqlExcLVGeral As String
Public Posicao As Integer
Public vPDO As Integer 'Variavel para armazenar ultimo numero de PDO criado
Public vCodModeloAval As Integer 'variavel do codigo do modelo de avaliação de eficacia usado na programação
Public vCodcoligada As Integer 'Variavel que armazena codigo da coligada ativa
Public vCaminhoAtu As String 'Variavel que armazena caminho + executál de atualização automática do SGCH

Public vControlaDim  As Integer 'Controla a quantidade de vezes q sera dimensionado o MeuLV
Public vsituacao As String 'armazena a situacao do colaborador apos a avaliacao do treinamento
Public vNota As String 'armazena a nota do colaborador/candidato apos a avaliacao do treinamento
Public vqtdava As String 'armazena a quantidade de questoes avaliadas

Public TiPo As String 'armazana o tipo de hospede (F/J) para exclusao
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
Public vIntegra As String  'Para informar se o SGCH esta integrado a outro sistema
Public vDadosTotvs(18) As String
Public colheDados(17) As String 'Guarda dados de importação de colaboradores de arquivo TXT

Public mStream As ADODB.Stream 'Para gravar imagem no Banco Totvs

Public vServerTotvs As String  'Armazena nome do servidor totvs
Public vBancoTotvs As String  'Armazena nome do banco totvs
Public vUsuBancoTovs As String  'Armazena usuario do banco totvs
Public vSenhaBancoTotvs As String  'Armazena senha do banco totvs

Public chamaForm As Form

Public MeuLV As New frmPesqGeral
Public NomeColLV(15) As String
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
Public vTabela1 As String, vTabela2 As String, vTabela3 As String, vTabela4 As String, vTabela5 As String, vTabela6 As String, vTabela7 As String, vTabela8 As String, vTabela9 As String, vTabela10 As String

'Public CodUsu As String ' codigo do usuário q esta logado
'Public NomUsu As String ' Nome do usuario
'Public CapturaCodigo As String ' Codigo da Empresa e do Contato
'Public Legenda As String ' Informa o significado (F)Fone (F) Fax (C) Celular
Public procnom As String, procnom1 As String
Public strAno As String 'Usada no relatorio de programação de cursos/treinamentos anual
Public vponteiro As Integer
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
    MsgBox "(ADO) Erro ao tentar acessar DB " & vbCrLf & Err.Number & " - Procure o administrador da rede " & Err.Description, 16, "Mensagem de erro"
    Exit Sub
End Sub

Public Sub Conexao()
'On Error GoTo Err1
    If sServerName = "" Then GoTo Err1
    Set cnBanco = New ADODB.Connection
    'If sSGBD = 1 Then
    '    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=False;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    'ElseIf sSGBD = 2 Then
        'cnBanco.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=True;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    
    
    
    cnBanco.Open "PROVIDER=MSDASQL.1; DRIVER={PostgreSQL UNICODE}; DATABASE=" & sDatabaseName & "; SERVER=" & sServerName & "; UID=" & sUsuName & "; PWD=" & sSenhaDB & "; ByteaAsLongVarBinary=1; Port=5432"
    
    
    'Else
    '    Resume Err1
    'End If
    frmSplash.Label5.Caption = "Conexão realizada com sucesso"
    Exit Sub
Err1:
    frmSplash.Label5.Caption = "Falha na conexão"
    MsgBox "Erro ao tentar acessar DB - Entre com as novas configurações do servidor "
    Exit Sub
End Sub

'ABAIXO CONEXÃO COM O BANCO DE DADOS RM
Public Function ConexaoTotvs()
On Error GoTo Err1
    Set cnBancoTotvs = New ADODB.Connection
    cnBancoTotvs.Open "Provider=SQLOLEDB.1;Password=" & vSenhaBancoTotvs & ";Persist Security Info=True;User ID=" & vUsuBancoTovs & ";Initial Catalog=" & vBancoTotvs & ";Data Source=" & vServerTotvs
    vIntegra = "S"
    achaSecaoSGCH
    criaTrigger
    Exit Function
Err1:
    MsgBox "Erro de conexão com Banco Totvs"
    vIntegra = "N"
    Exit Function
End Function

Public Function criaTrigger()
On Error GoTo Err
    'Essa rotina cria uma TRIGGER entre o Banco da TOTVS e o Banco do SGCH
    Dim sqlTrigger As String
    Dim rsTrigger As New ADODB.Recordset
    Dim rsVerTabSGCH As ADODB.Recordset
    
'Verifica se a tabela zSGCH_Demitidos existe no banco CORPORERM
    Set rsVerTabSGCH = cnBancoTotvs.OpenSchema(adSchemaTables, Array(Empty, Empty, "zSGCH_Demitidos", "Table"))
    If rsVerTabSGCH.EOF Then
        cnBancoTotvs.Execute "CREATE TABLE " & vBancoTotvs & ".dbo.zSGCH_Demitidos(" & _
        "chapa VARCHAR(10) NOT NULL," & _
        "controleSGCH CHAR(1) NOT NULL," & _
        "PRIMARY KEY (chapa))"
    End If
    rsVerTabSGCH.Close
'FIM TESTE
    
    sqlTrigger = "CREATE TRIGGER TriggerMonitoraTotvs on PFUNC for insert,update as Insert dbo.zSGCH_Demitidos " & _
                "Select CHAPA,'' from inserted Where CODSITUACAO = 'D'"
    
    'sqlTrigger = "CREATE TRIGGER [dbo].[TriggerMonitoraTotvs] on [dbo].[PFUNC]For insert,update as if (select count (*) from deleted) <> 0 " & _
    '            "Update B set B.ativo = 'N',B.datarecisao = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))),B.homologacaonum = 'Ver Totvs',b.homologacaoorgao = 'Ver Totvs'FROM dbo.PFUNC as A Inner join " & _
    '            sDatabaseName & ".dbo.tbcolaboradores as B on A.CHAPA=B.CODCOLABORADOR COLLATE SQL_Latin1_General_CP1_CI_AS Where A.CODSITUACAO = 'D'"
    rsTrigger.Open sqlTrigger, cnBancoTotvs
    Exit Function

Err:
    sqlTrigger = "ALTER TRIGGER [dbo].[TriggerMonitoraTotvs] on [dbo].[PFUNC] for insert,update as Insert dbo.zSGCH_Demitidos " & _
                "Select CHAPA,'' from inserted Where CODSITUACAO = 'D'"
    'sqlTrigger = "ALTER TRIGGER [dbo].[TriggerMonitoraTotvs] on [dbo].[PFUNC]For insert,update as if (select count (*) from deleted) <> 0 " & _
    '            "Update B set B.ativo = 'N',B.datarecisao = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))),B.homologacaonum = 'Ver Totvs',b.homologacaoorgao = 'Ver Totvs'FROM dbo.PFUNC as A Inner join " & _
    '            sDatabaseName & ".dbo.tbcolaboradores as B on A.CHAPA=B.CODCOLABORADOR COLLATE SQL_Latin1_General_CP1_CI_AS Where A.CODSITUACAO = 'D'"
    rsTrigger.Open sqlTrigger, cnBancoTotvs
End Function

Public Sub achaSecaoSGCH()
    Dim sqlSecaoSGCH As String
    Dim rsSecaoSGCH As New ADODB.Recordset
    Dim vIDSecao As Integer
    
    Dim sqlGravaSGCH As String
    Dim rsGravaSGCH As New ADODB.Recordset
    
    
    'sqlSecaoSGCH = "Select MAX(id)+1 from PSECAO"
    'rsSecaoSGCH.Open sqlSecaoSGCH, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    'vIDSecao = rsSecaoSGCH.Fields(0)
    'rsSecaoSGCH.Close
    'Set rsSecaoSGCH = Nothing
    
    'Cria seção de ADMISSÃO de colaboradores do SGCH
    sqlSecaoSGCH = "select * from PSECAO where DESCRICAO = 'SGCH'"
    rsSecaoSGCH.Open sqlSecaoSGCH, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    If rsSecaoSGCH.EOF Then
        sqlGravaSGCH = "Insert into " & _
                       "PSECAO(codcoligada,codigo,descricao,cgc,fpas,sat,rua,numero,bairro,estado,cidade,cep,pais,telefone,naoempregpropr,categoria,codterceirosinss," & _
                                "PERCENTTERCEIROS,percentacidtrab,proprantes5dia1,proprantes5dia2,centrantes5dia1,centrantes5dia2,CONTRIBSESIESENAI,distribpetroleo,pessoafisica," & _
                                "secaodesativada,identificacaocgc,enderecoalterou,codmunicipio,naturezajuridica,codcalendario,prefixoinscrfgts,primeiradeclcaged,encerramento," & _
                                "codfilial,coddepto,optasimples,alteracaocaged,codpagtogps,participapat,porteempresa,ddd,isentocontribsocial,vincpat5sal,vincpatmaior5sal,porcservprop," & _
                                "porcadmcozinha,porcrefeicaoconv,porcrefeicaotransp,porccestaalimento,PORCALIMCONVENIO,email,cnaerais,valorentidadesacumulado,idmemoambtrab,visivelorganograma," & _
                                "codigopai,reccreatedby,reccreatedon,recmodifiedby,recmodifiedon) " & _
                       "Values(1,'001.01.01.01','SGCH','19.431.980/0001-05','507','2511000','AV VITO GAGGIATO','SN','DISTRITO INDUSTRIAL','MG','SANTANA DO PARAISO', " & _
                                "'35167-000','BRASIL','3801-2600',3,'99','0079',5.80,3.00,0,0,0,0,0,0,0,0,1,0,'3158953','2062','0000001','01',1,2,1,'01',1,2,'2100', " & _
                                "0,2,'0031', 0,0,0,0,0,0,0,0,0,'pessoal@viga.ind.br','25110','0.0000',82,'T','001.01','mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ",'mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ")"
        rsGravaSGCH.Open sqlGravaSGCH, cnBancoTotvs
    End If
    rsSecaoSGCH.Close
    Set rsSecaoSGCH = Nothing
    
    'Cria seção de ALTERÇÃO FUNCIONAL de colaboradores do SGCH
    sqlSecaoSGCH = "select * from PSECAO where DESCRICAO = 'SGCH - Alteração funcional'"
    rsSecaoSGCH.Open sqlSecaoSGCH, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    If rsSecaoSGCH.EOF Then
        sqlGravaSGCH = "Insert into " & _
                       "PSECAO(codcoligada,codigo,descricao,cgc,fpas,sat,rua,numero,bairro,estado,cidade,cep,pais,telefone,naoempregpropr,categoria,codterceirosinss," & _
                                "PERCENTTERCEIROS,percentacidtrab,proprantes5dia1,proprantes5dia2,centrantes5dia1,centrantes5dia2,CONTRIBSESIESENAI,distribpetroleo,pessoafisica," & _
                                "secaodesativada,identificacaocgc,enderecoalterou,codmunicipio,naturezajuridica,codcalendario,prefixoinscrfgts,primeiradeclcaged,encerramento," & _
                                "codfilial,coddepto,optasimples,alteracaocaged,codpagtogps,participapat,porteempresa,ddd,isentocontribsocial,vincpat5sal,vincpatmaior5sal,porcservprop," & _
                                "porcadmcozinha,porcrefeicaoconv,porcrefeicaotransp,porccestaalimento,PORCALIMCONVENIO,email,cnaerais,valorentidadesacumulado,idmemoambtrab,visivelorganograma," & _
                                "codigopai,reccreatedby,reccreatedon,recmodifiedby,recmodifiedon) " & _
                       "Values(1,'001.01.01.02','SGCH - Alteração funcional','19.431.980/0001-05','507','2511000','AV VITO GAGGIATO','SN','DISTRITO INDUSTRIAL','MG','SANTANA DO PARAISO', " & _
                                "'35167-000','BRASIL','3801-2600',3,'99','0079',5.80,3.00,0,0,0,0,0,0,0,0,1,0,'3158953','2062','0000001','01',1,2,1,'01',1,2,'2100', " & _
                                "0,2,'0031', 0,0,0,0,0,0,0,0,0,'pessoal@viga.ind.br','25110','0.0000',82,'T','001.01','mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ",'mestre'," & Format(CStr(Date), "YYYY/MM/DD") & ")"
        rsGravaSGCH.Open sqlGravaSGCH, cnBancoTotvs
    End If
    rsSecaoSGCH.Close
    Set rsSecaoSGCH = Nothing
    
    
End Sub

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
    Else
        sql = "Select * from " & Tabela & " Order By " & Campo1
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
    'If Column Is Nothing Then
        For Each c In MeuLV.ListView1.ColumnHeaders
            Combo.AddItem c
        Next
        Combo.Text = Combo.List(0)
    'End If
End Sub

Public Sub CompoeComboNivel(Combo As ComboBox, txtBox As String)
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    sql = "select b.codnivel,b.nomenivel from tbtreinamentos as a inner join tbTreinamentosNiv as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento where a.codtreinamento = '" & Val(txtBox) & "'"
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
    Combo.Clear
    If Not rsTabela.EOF() Then
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            Combo.AddItem Format(rsTabela.Fields(0), "00") & " - " & rsTabela.Fields(1)
            Combo.ItemData(Combo.NewIndex) = Val(rsTabela.Fields(0))
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
    Variavel = Replace(campo, "%", "")
    RemoveMask = Variavel
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
    
    'CRIA BANCO SGCH
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDadosBanco(" & _
    "NomeServidor VARCHAR(50) NULL," & _
    "NomeBanco VARCHAR(50) NULL)"
    
    'TABELAS SGCH
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTreinamentos(" & _
    "codtreinamento NUMERIC NOT NULL," & _
    "nometreinamento VARCHAR(100) NOT NULL," & _
    "tipo VARCHAR(50) NOT NULL," & _
    "origem VARCHAR(30) NOT NULL," & _
    "conteudo TEXT NULL," & _
    "objetivo TEXT NULL," & _
    "introdutorio VARCHAR(1) NULL," & _
    "aplicavel VARCHAR(1) NULL," & _
    "tempoaplic VARCHAR(10) NULL," & _
    "mesanoaplic VARCHAR(10) NULL," & _
    "observacao TEXT NULL," & _
    "cargahoraria VARCHAR(30) NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "obrigatorio VARCHAR(1) NULL," & _
    "nivel VARCHAR(1) NULL," & _
    "valor FLOAT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codtreinamento,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbEscolaridade(" & _
    "codescolaridade NUMERIC NOT NULL," & _
    "nomeescolaridade VARCHAR(100) NULL," & _
    "peso FLOAT NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codescolaridade,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbHabilidades(" & _
    "codhabilidade NUMERIC NOT NULL," & _
    "nomehabilidade VARCHAR(100) NOT NULL," & _
    "peso NUMERIC NOT NULL," & _
    "descricao TEXT NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codhabilidade,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbSetores(" & _
    "codsetor NUMERIC NOT NULL," & _
    "nomesetor VARCHAR(100) NOT NULL," & _
    "coddepartamento NUMERIC NOT NULL," & _
    "descricao TEXT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codsetor,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbSetoresHistResp(" & _
    "codsetor NUMERIC NOT NULL," & _
    "coddepartamento NUMERIC NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "dataini DATETIME NOT NULL," & _
    "datafim DATETIME NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codsetor,coddepartamento,codcolaborador,codcoligada))"
        
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDepartamentos(" & _
    "coddepartamento NUMERIC NOT NULL," & _
    "nomedepartamento VARCHAR(100) NOT NULL," & _
    "descricao TEXT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (coddepartamento,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbDepartamentosHistResp(" & _
    "coddepartamento NUMERIC NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "dataini DATETIME NOT NULL," & _
    "datafim DATETIME NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (coddepartamento,codcolaborador,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbCargos(" & _
    "codcargo NUMERIC NOT NULL," & _
    "codcbo VARCHAR(30) NULL," & _
    "nomecargo VARCHAR(200) NOT NULL," & _
    "descricao TEXT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codcargo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMatriz(" & _
    "codmatriz NUMERIC NOT NULL," & _
    "coddepartamento NUMERIC NOT NULL," & _
    "codsetor NUMERIC NOT NULL," & _
    "codcargo NUMERIC NOT NULL," & _
    "nivel VARCHAR(5) NOT NULL," & _
    "atividades TEXT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "tempoMin VARCHAR(50) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmatriz,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTreinamentosRev(" & _
    "codtreinamento NUMERIC NOT NULL," & _
    "revisao VARCHAR(10) NOT NULL," & _
    "data DATETIME NOT NULL," & _
    "detalhes TEXT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codtreinamento,revisao,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTreinamentosInt(" & _
    "codtreinamento NUMERIC NOT NULL," & _
    "codsetor NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codtreinamento,codsetor,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTreinamentosObr(" & _
    "codtreinamento NUMERIC NOT NULL," & _
    "codsetor NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codtreinamento,codsetor,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTreinamentosNiv(" & _
    "codtreinamento NUMERIC NOT NULL," & _
    "codnivel NUMERIC NOT NULL," & _
    "nomenivel VARCHAR(100) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codtreinamento,codnivel,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMatrizEsc(" & _
    "codmatriz NUMERIC NOT NULL," & _
    "codescolaridade NUMERIC NOT NULL," & _
    "pontuacao FLOAT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmatriz,codescolaridade,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMatrizExp(" & _
    "codmatriz NUMERIC NOT NULL," & _
    "codcargo NUMERIC NOT NULL," & _
    "tmpoexp VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmatriz,codcargo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMatrizCur(" & _
    "codmatriz NUMERIC NOT NULL," & _
    "codtreinamento NUMERIC NOT NULL," & _
    "codnivel NUMERIC NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmatriz,codtreinamento,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMatrizHab(" & _
    "codmatriz NUMERIC NOT NULL," & _
    "codhabilidade NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmatriz,codhabilidade,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMatrizRev(" & _
    "codmatriz NUMERIC NOT NULL," & _
    "revisao VARCHAR(10) NOT NULL," & _
    "data DATETIME NOT NULL," & _
    "detalhes TEXT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmatriz,revisao,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradoresExp(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "tipo VARCHAR(30) NOT NULL," & _
    "nomeempresa VARCHAR(100) NOT NULL," & _
    "codcargo NUMERIC NOT NULL," & _
    "tempoexp VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (cpf,nomeempresa,codcargo,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbPrintMatriz(" & _
    "campo1 NUMERIC NULL," & _
    "campo2 NUMERIC NULL," & _
    "campo3 VARCHAR(100) NULL," & _
    "campo4 VARCHAR(100) NULL," & _
    "campo5 VARCHAR(100) NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbPrintHFunc(" & _
    "campo1 VARCHAR(200) NULL," & _
    "campo2 VARCHAR(200) NULL," & _
    "campo3 VARCHAR(200) NULL," & _
    "campo4 VARCHAR(200) NULL," & _
    "campo5 VARCHAR(200) NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradoresHab(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "tipo VARCHAR(30) NOT NULL," & _
    "codhabilidade NUMERIC NOT NULL," & _
    "pontuacao FLOAT NULL," & _
    "codmatriz NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (cpf,codhabilidade,codmatriz,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradoresCur(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "tipo VARCHAR(30) NOT NULL," & _
    "codtreinamento NUMERIC NOT NULL," & _
    "origem VARCHAR(2) NOT NULL," & _
    "ID INT NOT NULL IDENTITY," & _
    "codnivel NUMERIC NULL," & _
    "codcoligada INT NOT NULL," & _
    "datacur DATETIME NULL," & _
    "nometreinamento VARCHAR(300) NULL," & _
    "PRIMARY KEY (ID,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradoresEsc(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "tipo VARCHAR(30) NOT NULL," & _
    "codescolaridade NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (cpf,codescolaridade,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradoresHist(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "codmatriz NUMERIC NOT NULL," & _
    "data DATETIME NOT NULL," & _
    "motivo TEXT NULL," & _
    "observacao TEXT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "sequencia NUMERIC NOT NULL," & _
    "tipo VARCHAR(50) NOT NULL," & _
    "codrequisicao NUMERIC NULL," & _
    "datasai DATETIME NULL," & _
    "justificativa VARCHAR(300) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (cpf,codmatriz,data,tipo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradores(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL,datacadastro DATETIME NOT NULL," & _
    "nomecolaborador VARCHAR(100) NOT NULL,datanascimento DATETIME NULL," & _
    "sexo VARCHAR(10) NULL," & _
    "estadocivil VARCHAR(30) NULL," & _
    "nacionalidade VARCHAR(50) NULL," & _
    "naturalidade VARCHAR(50) NULL," & _
    "ufnaturalidade VARCHAR(2) NULL," & _
    "ctpsnumero VARCHAR(50) NULL," & _
    "ctpsserie VARCHAR(30) NULL," & _
    "cnhnumero VARCHAR(50) NULL," & _
    "cnhtipo VARCHAR(30) NULL, datarecisao DATETIME NULL," & _
    "homologacaonum VARCHAR(50) NULL,homologacaoorgao VARCHAR(100) NULL," & _
    "ativo VARCHAR(1) NULL, mediageral FLOAT NULL," & _
    "foto TEXT NULL, observacao TEXT NULL," & _
    "compav VARCHAR(10) NULL, email VARCHAR(100) NULL," & _
    "tipo VARCHAR(30) NULL, telefone VARCHAR(30) NULL," & _
    "celular VARCHAR(30) NULL, codrequisicao NUMERIC NULL," & _
    "geroupen VARCHAR(1) NULL, obsadm VARCHAR(200) NULL," & _
    "id INT NOT NULL IDENTITY, autorizacao NUMERIC NULL," & _
    "codcoligada INT NOT NULL," & _
    "dataafastamento DATETIME NULL," & _
    "PRIMARY KEY (cpf,codcolaborador,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbRequisicoes(" & _
    "codrequisicao NUMERIC NOT NULL," & _
    "datarequisicao DATETIME NOT NULL," & _
    "tipo VARCHAR(30) NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "nomerequisitante VARCHAR(50) NULL," & _
    "departamentorequisitante VARCHAR(50) NULL," & _
    "setorrequisitante VARCHAR(50) NOT NULL," & _
    "origem VARCHAR(10) NOT NULL," & _
    "nomeempresa VARCHAR(50) NULL," & _
    "ativo VARCHAR(1) NOT NULL," & _
    "observacao TEXT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codrequisicao,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbRequisicoesCargos(" & _
    "codrequisicao NUMERIC NOT NULL," & _
    "codmatriz NUMERIC NOT NULL," & _
    "numvagas NUMERIC NOT NULL," & _
    "dataprevisaoadm DATETIME NOT NULL," & _
    "motivo TEXT NULL," & _
    "observacao TEXT NULL," & _
    "qtdcolaboradores NUMERIC NOT NULL," & _
    "qtdocupada NUMERIC NULL," & _
    "status VARCHAR(30) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codrequisicao,codmatriz,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbRequisicoesAprovadores(" & _
    "codrequisicao NUMERIC NOT NULL," & _
    "tipo VARCHAR(50) NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "nomeaprovador VARCHAR(50) NOT NULL," & _
    "sequencia NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codrequisicao,sequencia,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbPendentesCur(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "codmatriz NUMERIC NOT NULL," & _
    "codtreinamento NUMERIC NOT NULL," & _
    "codprogramacao NUMERIC NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "id NUMERIC NOT NULL," & _
    "status VARCHAR(20) NOT NULL," & _
    "tipoprogramacao NUMERIC NOT NULL," & _
    "situacao VARCHAR(50) NULL," & _
    "nota FLOAT NULL," & _
    "observacao TEXT NULL," & _
    "obsavaliacao TEXT NULL," & _
    "codnivel NUMERIC NULL," & _
    "codINTD NUMERIC NULL," & _
    "codcoligada INT NOT NULL," & _
    "cargoorigem VARCHAR(100) NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAvaliacao(" & _
    "codavaliacao NUMERIC NOT NULL," & _
    "nomeavaliacao TEXT NOT NULL," & _
    "tipo VARCHAR(2) NOT NULL," & _
    "peso NUMERIC NOT NULL," & _
    "ativo VARCHAR(1) NOT NULL," & _
    "descricao TEXT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codavaliacao,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAvaliacaoTrei(" & _
    "codprogramacao NUMERIC NOT NULL," & _
    "CPF VARCHAR(50) NOT NULL," & _
    "codavaliacao NUMERIC NOT NULL," & _
    "pontuacao FLOAT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprogramacao,CPF,codavaliacao,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAvaliacaoProg(" & _
    "codavaliacao NUMERIC NOT NULL," & _
    "codmodelo INT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codavaliacao,codmodelo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProgramacao(" & _
    "codprogramacao NUMERIC NOT NULL," & _
    "dataprogramacao DATETIME NOT NULL," & _
    "entidade VARCHAR(50) NOT NULL," & _
    "local VARCHAR(100) NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "datainicio DATETIME NOT NULL,datafim DATETIME NOT NULL," & _
    "horainicio DATETIME NOT NULL,horafim DATETIME NOT NULL," & _
    "dae BIT NOT NULL," & _
    "metodo NUMERIC NULL," & _
    "metodooutro VARCHAR(50) NULL," & _
    "nota FLOAT NULL," & _
    "observacao TEXT NULL," & _
    "status VARCHAR(30) NOT NULL," & _
    "ativo VARCHAR(1) NOT NULL," & _
    "avaltipo VARCHAR(50) NULL," & _
    "avalnome VARCHAR(50) NULL," & _
    "avaldata DATETIME NULL," & _
    "codmodelo INT NULL," & _
    "metodoA BIT NULL," & _
    "metodoT BIT NULL," & _
    "metodoS BIT NULL," & _
    "metodoPT BIT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprogramacao,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProgramacaoInstrutores(" & _
    "codprogramacao NUMERIC NOT NULL," & _
    "origem VARCHAR(20) NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "nomeinstrutor VARCHAR(50) NOT NULL," & _
    "tipoaula VARCHAR(30) NOT NULL," & _
    "sequencia NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprogramacao,sequencia,codcoligada))"
'-----
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProcessos(" & _
    "codprocesso NUMERIC NOT NULL," & _
    "codrequisicao NUMERIC NOT NULL," & _
    "datainicio DATETIME NOT NULL," & _
    "datafim DATETIME NOT NULL," & _
    "listar VARCHAR(1) NOT NULL," & _
    "linhas NUMERIC NOT NULL," & _
    "status VARCHAR(30) NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprocesso,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProcessosCargos(" & _
    "codprocesso NUMERIC NOT NULL," & _
    "codmatriz NUMERIC NOT NULL," & _
    "numvagas NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprocesso,codmatriz,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProcessoListaTmp(" & _
    "cpf VARCHAR(50) NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "matrizcpf VARCHAR(50) NOT NULL," & _
    "cargocpf VARCHAR(50) NOT NULL," & _
    "tipo VARCHAR(50) NOT NULL," & _
    "cargopesq VARCHAR(50) NOT NULL," & _
    "nota VARCHAR(50) NOT NULL," & _
    "matrizpesq VARCHAR(50) NOT NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProcessosParticipantes(" & _
    "codprocesso NUMERIC NOT NULL," & _
    "CPF VARCHAR(50) NOT NULL," & _
    "matrizpesq NUMERIC NOT NULL," & _
    "tipo VARCHAR(50) NOT NULL," & _
    "matrizcargo NUMERIC NOT NULL," & _
    "nota FLOAT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprocesso,CPF,matrizpesq,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbProcessosAdm(" & _
    "codprocesso NUMERIC NOT NULL," & _
    "CPF VARCHAR(50) NOT NULL," & _
    "matrizpesq NUMERIC NOT NULL," & _
    "tipo VARCHAR(30) NOT NULL," & _
    "matrizcargo NUMERIC NOT NULL," & _
    "nota FLOAT NOT NULL," & _
    "observacao TEXT NULL," & _
    "codcolaborador VARCHAR(50) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codprocesso,CPF,matrizpesq,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbINTD(" & _
    "codINTD NUMERIC NOT NULL," & _
    "datainicio DATETIME NOT NULL," & _
    "datafim DATETIME NOT NULL," & _
    "tipoINTD VARCHAR(1) NOT NULL," & _
    "tiposolicitante VARCHAR(30) NOT NULL," & _
    "codsolicitante NUMERIC NOT NULL," & _
    "nomesolicitante VARCHAR(50) NOT NULL," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "codmatriz NUMERIC NULL," & _
    "status VARCHAR(30) NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "objetivo TEXT NULL," & _
    "mediageral FLOAT NULL," & _
    "resultado VARCHAR(10) NULL," & _
    "observacao TEXT NULL," & _
    "codcoligada INT NOT NULL," & _
    "mediaescolar FLOAT NULL," & _
    "cargoOrigem VARCHAR(100) NULL," & _
    "PRIMARY KEY (codINTD,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbINTDCur(" & _
    "codINTD NUMERIC NOT NULL," & _
    "codTreinamento NUMERIC NOT NULL," & _
    "codnivel NUMERIC NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codINTD,codTreinamento,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbINTDHab(" & _
    "codINTD NUMERIC NOT NULL," & _
    "codHabilidade NUMERIC NOT NULL," & _
    "pontuacao FLOAT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codINTD,codHabilidade,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAutorizacao(" & _
    "id INT NOT NULL IDENTITY," & _
    "cpf VARCHAR(20) NOT NULL," & _
    "tipo VARCHAR(20) NOT NULL," & _
    "nota VARCHAR(20) NOT NULL," & _
    "solicitacao VARCHAR(300) NOT NULL," & _
    "observacao TEXT NULL," & _
    "status CHAR(1) NULL," & _
    "datasolicitacao VARCHAR(20) NULL," & _
    "solicitante VARCHAR(50) NOT NULL," & _
    "aprovador VARCHAR(30) NULL," & _
    "datadecisao DATETIME NULL," & _
    "decisao VARCHAR(20) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbModeloProg(" & _
    "codmodelo INT NOT NULL," & _
    "nomemodelo VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmodelo,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfCertificado(" & _
    "id INT NOT NULL IDENTITY,Memo1 TEXT NULL," & _
    "certificadora VARCHAR(100) NULL,titulo VARCHAR(50) NULL," & _
    "logo CHAR(1) NULL,borda CHAR(1) NULL," & _
    "fundo CHAR(1) NULL,fundocaminho VARCHAR(100) NULL," & _
    "fontCab VARCHAR(50) NULL,fontCorp VARCHAR(50) NULL," & _
    "fontRod VARCHAR(50) NULL,fontCert VARCHAR(50) NULL," & _
    "tamFontCab VARCHAR(5) NULL," & _
    "tamFontCorp VARCHAR(5) NULL," & _
    "tamFontRod VARCHAR(5) NULL," & _
    "tamFontCert VARCHAR(5) NULL," & _
    "alinFontCorp NUMERIC NULL," & _
    "alinFontRod NUMERIC NULL," & _
    "alinFontCab NUMERIC NULL," & _
    "alinFontCer NUMERIC NULL," & _
    "nometreinamento VARCHAR(50) NULL," & _
    "datainicio VARCHAR(15) NULL," & _
    "datafim VARCHAR(15) NULL," & _
    "cargahoraria VARCHAR(50) NULL," & _
    "responsavel VARCHAR(50) NULL," & _
    "dataemissao VARCHAR(15) NULL," & _
    "logocaminho VARCHAR(100) NULL," & _
    "titResp VARCHAR(50) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColabCertificado(" & _
    "id INT NOT NULL," & _
    "nomecolaborador VARCHAR(100) NOT NULL," & _
    "nota VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,nomecolaborador,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbAvaliacaoDesempenho(" & _
    "id INT NOT NULL," & _
    "dias NUMERIC NOT NULL," & _
    "tipo VARCHAR(20) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbABS(" & _
    "id INT NOT NULL," & _
    "tipo VARCHAR(20) NOT NULL," & _
    "oc1 NUMERIC NOT NULL," & _
    "oc2 NUMERIC NOT NULL," & _
    "pontos NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbListaADP(" & _
    "id INT NOT NULL IDENTITY," & _
    "codcolaborador VARCHAR(50) NOT NULL," & _
    "tipoADP VARCHAR(30) NOT NULL,dias NUMERIC NOT NULL," & _
    "dataavaliacao DATETIME NULL,datavencimento DATETIME NOT NULL," & _
    "datadevolucao DATETIME NOT NULL,codrespADP VARCHAR(50) NULL," & _
    "nomerespADP VARCHAR(50) NULL,ausenciaano NUMERIC NULL," & _
    "atrasoano NUMERIC NULL,codrespABS VARCHAR(50) NULL," & _
    "nomerespABS VARCHAR(50) NULL," & _
    "observacao TEXT NULL," & _
    "indicacaotipo NUMERIC NULL," & _
    "indicacaomod1 NUMERIC NULL," & _
    "indicacaomod2 NUMERIC NULL," & _
    "indicacaomod3 NUMERIC NULL," & _
    "indicacaomod4 NUMERIC NULL," & _
    "indicacaomod5 NUMERIC NULL," & _
    "indicacaomod6 NUMERIC NULL," & _
    "indicacaooutros VARCHAR(50) NULL," & _
    "statusimpressao INT NULL," & _
    "statusavaliacao VARCHAR(30) NULL," & _
    "ativo CHAR(1) NULL," & _
    "nota FLOAT NULL," & _
    "codcoligada INT NOT NULL," & _
    "cargoADP VARCHAR(100) NULL," & _
    "PRIMARY KEY (id,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbListaADPItens(" & _
    "idADP INT NOT NULL," & _
    "codavaliacao NUMERIC NOT NULL," & _
    "nota FLOAT NOT NULL," & _
    "dimensao VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (idADP,codavaliacao,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbModeloADP(" & _
    "codmodelo INT NOT NULL," & _
    "nomemodelo VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codmodelo,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbModeloADPItens(" & _
    "codavaliacao NUMERIC NOT NULL," & _
    "codmodelo INT NOT NULL," & _
    "dimensao VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codavaliacao,codmodelo,codcoligada))"

'-----
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
    "afastdias NUMERIC NULL," & _
    "afasttreiint VARCHAR(1) NULL," & _
    "afasttreiobr VARCHAR(1) NULL," & _
    "PRIMARY KEY (mediaaprovacao,codcoligada))"
    
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
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfConvocacao(" & _
    "id INT NOT NULL IDENTITY," & _
    "tipoconvocacao NUMERIC NOT NULL," & _
    "tipotreinamento VARCHAR(100) NOT NULL," & _
    "responsavel VARCHAR(50) NOT NULL," & _
    "texto TEXT NOT NULL," & _
    "dataconvocacao DATETIME NOT NULL," & _
    "horarioini DATETIME NOT NULL," & _
    "horariofim DATETIME NOT NULL," & _
    "cargahoraria VARCHAR(30) NOT NULL," & _
    "local VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
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
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbColaboradoresIntTotvs(" & _
    "id INT NOT NULL," & _
    "modulo VARCHAR(50) NOT NULL," & _
    "sexo VARCHAR(50) NOT NULL," & _
    "grauinst VARCHAR(50) NOT NULL," & _
    "tipoadm VARCHAR(50) NOT NULL," & _
    "motadm VARCHAR(50) NOT NULL," & _
    "forreceb VARCHAR(50) NOT NULL," & _
    "situacao VARCHAR(50) NOT NULL," & _
    "tipofunc VARCHAR(50) NOT NULL," & _
    "hortrab VARCHAR(50) NOT NULL," & _
    "funcao VARCHAR(50) NOT NULL," & _
    "secao VARCHAR(50) NOT NULL," & _
    "contsind VARCHAR(50) NOT NULL," & _
    "rais VARCHAR(50) NOT NULL," & _
    "memsind VARCHAR(50) NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbTreinamentosAgr(" & _
    "codigoTrei NUMERIC NOT NULL," & _
    "codigoTreiGrup NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigoTrei,codigoTreiGrup,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfLV(" & _
    "nmusuario VARCHAR(50) NULL," & _
    "idmodulo NUMERIC NOT NULL," & _
    "indice NUMERIC NOT NULL," & _
    "posicao NUMERIC NOT NULL," & _
    "largura FLOAT NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "PRIMARY KEY (id))"
    
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
    
    
    'CRIA TABELAS ADMINISTRATIVAS
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbConfGrupo(" & _
    "idgrupo NUMERIC NOT NULL," & _
    "idmenu NUMERIC NOT NULL," & _
    "idsub NUMERIC NOT NULL," & _
    "tipo VARCHAR(20) NOT NULL," & _
    "nome VARCHAR(50) NOT NULL," & _
    "status VARCHAR(1) NOT NULL," & _
    "id INT NOT NULL IDENTITY," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (id,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbSenha(" & _
    "usuario VARCHAR(50) NOT NULL," & _
    "senha VARCHAR(50) NOT NULL," & _
    "codigo NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"
    
    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMenu(" & _
    "idmenu NUMERIC NULL," & _
    "idsub NUMERIC NULL," & _
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

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbUsuMultiplic(" & _
    "codusuario NUMERIC NOT NULL," & _
    "codtreinamento NUMERIC NOT NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codusuario,codtreinamento,codcoligada))"

    oConn.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbGrupo(" & _
    "codigo NUMERIC NOT NULL," & _
    "descricao VARCHAR(50) NOT NULL," & _
    "ativo VARCHAR(1) NULL," & _
    "codcoligada INT NOT NULL," & _
    "PRIMARY KEY (codigo,codcoligada))"
    
    'ABAIXO: CRIA CONFIGURAÇÃO PARA USUÁRIO ADMINISTRADOR
    oConn.Close
    'oConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    
    vCodcoligada = 1 'Primeiro cadastro de coligada
    
    'If sSGBD = 1 Then
        'oConn.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=False;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    'ElseIf sSGBD = 2 Then
        oConn.Open "Provider=SQLOLEDB.1;Password=" & sSenhaDB & ";Persist Security Info=True;User ID=" & sUsuName & ";Initial Catalog=" & sDatabaseName & ";Data Source=" & sServerName
    'End If

    SqlSenha = "Insert into tbSenha(usuario,senha,codigo,codcoligada) Values('adm','123',1,'" & vCodcoligada & "');"
    rsSenha.Open SqlSenha, oConn
    
    SqlUsuario = "Insert into tbUsuarios(codigo,nome,codgrupo,ativo,codcoligada) Values(1,'Administrador do sistema',1,'S','" & vCodcoligada & "');"
    rsUsuario.Open SqlUsuario, oConn
    
    SqlGrupo = "Insert into tbGrupo(codigo,descricao,ativo,codcoligada) Values(1,'Administradores','S','" & vCodcoligada & "');"
    rsGrupo.Open SqlGrupo, oConn
    
    
    SqlConfGrupo = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,1,'TAB','Cadastros','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,1,'CAT','Colaboradores','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,2,'CAT','Candidatos','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,3,'CAT','Departamentos','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,4,'CAT','Setores','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,5,'CAT','Cargos','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,6,'CAT','Habilidades funcionais','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,7,'CAT','Formação escolar','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,1,8,'CAT','Avaliação do treinamento','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,2,1,'TAB','Recrutamento','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,2,1,'CAT','Requisição de pessoal','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,2,2,'CAT','Processo seletivo','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,1,'TAB','Capacitação','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,1,'CAT','Cursos/treinamentos','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,2,'CAT','Matriz de capacitação','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,3,'CAT',' INTD ','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,4,'CAT','Programação','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,5,'CAT','Restrições','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,3,6,'CAT',' ADP ','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,4,1,'TAB','Configurações','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,4,1,'CAT','Usuários','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,4,2,'CAT','Grupos','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,4,3,'CAT','Sistema','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,5,1,'TAB','Sobre','S'," & vCodcoligada & ");" & _
                "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,5,1,'CAT','Sobre SGCH','S'," & vCodcoligada & ");Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada) Values(1,5,2,'CAT','Ajuda SGCH','S'," & vCodcoligada & ");"
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
       
    MsgBox "Tabelas criadas com sucesso", vbInformation, "SGCH"
    Exit Function
Err1:
    'MsgBox "(ADO) Erro ao criar Tabela de dados: " & vbCrLf & Err.Number & " - Tabela já Existe - " & Err.Description, 16, "Mensagem de erro"
    Resume Next
    'Exit Function
End Function

'Public Function DesabBotoesN0()
'    Dim X As Integer
'    For X = 0 To frmMenu2.chamCad.Count - 1
'        frmMenu2.chamCad(X).UseGreyscale = True
'    Next
'End Function

Public Function DesabBotoesN1(frm As Form)
    Dim X As Integer
    For X = 0 To frm.cmdconsulta.Count - 1
        frm.cmdconsulta(X).UseGreyscale = True
    Next
End Function

'Public Function HabBotoesN0()
'    Dim X As Integer
'    For X = 0 To frmMenu2.chamCad.Count - 1
'        frmMenu2.chamCad(X).UseGreyscale = False
'    Next
'End Function

Public Function HabBotoesN1(frm As Form)
    Dim X As Integer
    For X = 0 To frm.cmdconsulta.Count - 1
        frm.cmdconsulta(X).UseGreyscale = False
    Next
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'ROTINAS/FUNÇÕES DO LISTVIEW GENERICO - DAKI PARA BAIXO
Public Function MontaFiltro()
    If Formulario = "Cargos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Departamentos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Setores" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Treinamentos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Habilidades" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Matriz" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Escolaridade" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Colaboradores" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Demitidos"
        frmFiltro.Combo1.List(3) = "Afastados"
        frmFiltro.Combo1.List(4) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Candidatos" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Requisição" Then
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
    If Formulario = "Avaliação" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "Reprovados" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
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
    If Formulario = "Processo Seletivo" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "INTD" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
    If Formulario = "PDO" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Avaliados"
        frmFiltro.Combo1.List(1) = "Não Avaliados"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Não Avaliados"
    End If
    If Formulario = "ADP" Then
        'TiPo = False
        frmFiltro.Combo1.List(0) = "Ativos"
        frmFiltro.Combo1.List(1) = "Não ativos"
        frmFiltro.Combo1.List(2) = "Todos"
        frmFiltro.Combo1.Text = "Ativos"
    End If
End Function

Public Function MontaCabLV(Cab0 As String, Cab1 As String, Cab2 As String, Cab3 As String, Cab4 As String, Cab5 As String, Cab6 As String, Cab7 As String, Cab8 As String, Cab9 As String, Cab10 As String, Cab11 As String, Cab12 As String, Cab13 As String, Cab14 As String, Cab15 As String)
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
End Function

Public Function DimensionaLV(NomeLV As String)
    MeuLV.Frame1.Caption = NomeLV
    'frmMenu2.StatusBar1.Panels(3) = Legenda
    MeuLV.Top = frmMenu2.ACPRibbon1.Height + 10
    MeuLV.Left = frmMenu2.Left + 130
    MeuLV.Width = frmMenu2.Width - 300

    MeuLV.Frame1.Width = MeuLV.Width - (MeuLV.Width * 1.5 / 100)
    MeuLV.ListView1.Width = MeuLV.Frame1.Width - (MeuLV.Frame1.Width * 1.5 / 100)

    MeuLV.Height = frmMenu2.Height - 2700
    MeuLV.Frame1.Height = MeuLV.Height - 250
'    MeuLV.ListView1.Height = MeuLV.Frame1.Height - (MeuLV.Frame1.Height * 15 / 90)
    MeuLV.ListView1.Height = MeuLV.Frame1.Height - 1200 '(MeuLV.Frame1.Height * 15 / 90)

    Set DimensionaLV = Nothing
End Function

Public Function MontaCabecalhoLV()
    Dim X As Integer
    'Limpa o cabeçalho antes de compor novamente
    MeuLV.ListView1.ColumnHeaders.Clear
    With MeuLV.ListView1.ColumnHeaders
        For X = 0 To 15
            If NomeColLV(X) = "" Then Exit Function
            .Add , , NomeColLV(X), Len(NomeColLV(X)) * 144
            QtdColReal = QtdColReal + 1
        Next
    End With
    
    Set MontaCabecalhoLV = Nothing
    
End Function

Public Function MontaDadosLV(ZeroPriCol As String)
On Error GoTo Err
    ' Declaração de variaveis
    Dim rsListview As New ADODB.Recordset ' Variavel que vai receber os dados da tabela
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim X As Integer, Y As Integer
    rsListview.Open SqlLV, cnBanco, adOpenKeyset, adLockReadOnly
    MeuLV.ListView1.ListItems.Clear 'Limpa o listview
    If rsListview.RecordCount <> 0 Then frmMenu2.ProgressBar1.Max = rsListview.RecordCount
    X = 0
    While Not rsListview.EOF
        Y = 0
        frmMenu2.ProgressBar1.Value = X
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
    frmMenu2.ProgressBar1.Value = 0
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
    'MeuLV.ListView1.Sorted = True
    'MeuLV.ListView1.SortKey = 0
    'MeuLV.ListView1.SortOrder = lvwAscending
    rsListview.Close
    Set rsListview = Nothing
    frmMenu2.StatusBar1.Panels(5).Text = "Registros: " & MeuLV.ListView1.ListItems.Count
    
    Set MontaDadosLV = Nothing
    
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
        frmMenu2.ProgressBar1.Value = X
        Set ItemLst = MeuLV.ListView1.ListItems.Item(X)
        'NEGRITO NOS ITENS DA COLUNA
        If negritoCol = "S" Then ItemLst.ListSubItems(posCol).Bold = True
        'COR VERDE/VERMELHO NOS ITENS DA COLUNA
        If corCol = "S" Then
            If Formulario = "ADP" Then
                If Date > CDate(ItemLst.ListSubItems(5)) And ItemLst.ListSubItems(9) <> "Concluido" Then
                    'Linhas abaixo exibe ADP vencidas através de cores
                    ItemLst.ForeColor = &HFF&
                    ItemLst.ListSubItems(1).ForeColor = &HFF&
                    ItemLst.ListSubItems(2).ForeColor = &HFF&
                    ItemLst.ListSubItems(3).ForeColor = &HFF&
                    ItemLst.ListSubItems(4).ForeColor = &HFF&
                    ItemLst.ListSubItems(5).ForeColor = &HFF&
                    ItemLst.ListSubItems(6).ForeColor = &HFF&
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
            If ItemLst.SubItems(posCol) = "S" Then
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "OK"
            Else
                ItemLst.ListSubItems.Item(posCol).ReportIcon = "EXC"
            End If
            ItemLst.SubItems(posCol) = ""
        End If
        'ALINHAMENTO DA COLUNA
        If alinhaCol = "D" Then MeuLV.ListView1.ColumnHeaders(posCol + 1).Alignment = lvwColumnRight Else MeuLV.ListView1.ColumnHeaders(posCol + 1).Alignment = lvwColumnLeft
    Next
    frmMenu2.ProgressBar1.Value = 0
    Legenda = ""
    
    Set PersonaColLV = Nothing
    
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Function

Public Sub ExcluirDadosLV()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim rsExcLVGeral As New ADODB.Recordset
    cnBanco.BeginTrans
    If MsgBox("Confirma exclusão do " & LegendaExc & " selecionado?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
        'SqlExcLVGeral = "Delete from tbHabilidades where codHabilidade= " & Val(varGlobal)
        'rsExcLVGeral.Open SqlExcLVGeral, cnBanco
        MsgBox "Registro excluido com sucesso", vbInformation, "Ok!"
        'rsExcLVGeral.Update
    End If
    cnBanco.CommitTrans
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Public Sub ExcluirItemLV(LV As ListView)
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

'ROTINAS/FUNÇÕES DO LISTVIEW GENERICO - DAKI PARA CIMA
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Avaliador(TiPo As String)
'On Error GoTo Err
'On Error Resume Next
    Dim rsAvaliador As New ADODB.Recordset
    Dim SqlAvaliador As String
    Dim X As Integer
    Dim CONTADOR As Double, ConverTido As Double
    
    chamaForm.mskCadMatriz.PromptInclude = False
    
    Dim PontosColabExp As Double
    Dim PontosTotaisHab As Double
    Dim PontosTotaisTrei As Double
    Dim PontosTotaisFor As Double
    CONTADOR = 0
    
    If chamaForm.Caption = "Cadastro de colaboradores" Then
        For X = 0 To 4
            If chamaForm.chkAvaliador(X).Value = 1 Then
                CONTADOR = CONTADOR + 1
            End If
        Next
    Else
        For X = 0 To 3
            If chamaForm.chkAvaliador(X).Value = 1 Then
                CONTADOR = CONTADOR + 1
            End If
        Next
    End If
    chamaForm.Label37 = ""
    chamaForm.Label38 = ""
    chamaForm.Label39 = ""
    chamaForm.Label40 = ""
    chamaForm.Label41 = ""
    If chamaForm.Caption = "Cadastro de colaboradores" Then
        chamaForm.Label43 = ""
    End If
    If CONTADOR = 0 Then Exit Function
    'PRIMEIRO CHECKBOX - EXPERIENCIA
    If chamaForm.chkAvaliador(0).Value = 1 Then
        Dim PontosMatrizExp As Double
        Dim ContCargoMatExp As Double
        
        SqlAvaliador = "select * from tbMatrizExp as a left join tbColaboradoresExp as b on a.codcargo = b.codcargo and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and b.cpf = '" & chamaForm.mskCadMatriz & "' and     b.tipo = '" & TiPo & "' where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
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
        SqlAvaliador = "select a.cpf,a.codmatriz,a.codhabilidade,a.pontuacao,b.peso from tbColaboradoresHab as a inner join tbHabilidades as b on a.codcoligada = '" & vCodcoligada & "' and a.codhabilidade = b.codhabilidade and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and a.cpf = '" & chamaForm.mskCadMatriz & "' and     a.tipo = '" & TiPo & "' where a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        PontosMatrizHab = 0
        PontosColabHab = 0
        While Not rsAvaliador.EOF
            PontosMatrizHab = PontosMatrizHab + rsAvaliador.Fields(4)
            PontosColabHab = PontosColabHab + rsAvaliador.Fields(3)
            rsAvaliador.MoveNext
        Wend
        If PontosColabHab > 0 Then PontosTotaisHab = (PontosColabHab * 100) / PontosMatrizHab
        If PontosMatrizHab <= 0 Then PontosTotaisHab = 100
        
        If PontosTotaisHab >= MediaGlobal Then
            chamaForm.Label38.ForeColor = &H8000&
        ElseIf PontosTotaisHab < MediaGlobal And PontosTotaisHab >= vAprovadoRest Then
            chamaForm.Label38.ForeColor = &H80FF&
        Else
            chamaForm.Label38.ForeColor = &HC0&
        End If
        
        
        chamaForm.Label38 = Format(PontosTotaisHab, "#,##0.00;(#,##0.00)") & " %"
        rsAvaliador.Close
        Set rsAvaliador = Nothing
    End If
    'TERCEIRO CHECKBOX - CURSOS/TREINAMENTOS
    If chamaForm.chkAvaliador(2).Value = 1 Then
        Dim PontosMatrizTrei As Double
        Dim PontosColabTrei As Double
        
'        SqlAvaliador = "select * from tbMatrizCur as a left join tbcolaboradoresCur as b on a.codtreinamento = b.codtreinamento and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & TiPo & "' where a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        SqlAvaliador = "select * from tbMatrizCur as a left join tbcolaboradoresCur as b on a.codtreinamento = b.codtreinamento and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & TiPo & "' and b.codnivel >= a.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
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
        Dim PesoColab As Integer, PesoMatriz As Integer, pontoMaior As Integer
        
        SqlAvaliador = "select c.codmatriz,c.codescolaridade,c.pontuacao,b.cpf,b.tipo,b.codescolaridade,a.peso from tbescolaridade as a left join tbcolaboradoresesc as b on a.codescolaridade = b.codescolaridade and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & TiPo & "' left join tbmatrizEsc as c on a.codescolaridade = c.codescolaridade and c.codmatriz = '" & Val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' where a.codcoligada = '" & vCodcoligada & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        PontosColabFor = 0
        VerificaNull = 0
        pontoMaior = 0
        Do While Not rsAvaliador.EOF
            If Not IsNull(rsAvaliador.Fields(5)) And rsAvaliador.Fields(6) > PesoColab Then PesoColab = rsAvaliador.Fields(6)
            If Not IsNull(rsAvaliador.Fields(1)) And rsAvaliador.Fields(6) > PesoMatriz Then PesoMatriz = rsAvaliador.Fields(6)
            If Not IsNull(rsAvaliador.Fields(2)) Then PontosColabFor = rsAvaliador.Fields(2)
            If Not IsNull((rsAvaliador.Fields(2))) And Not IsNull(rsAvaliador.Fields(3)) Then
                pontoMaior = rsAvaliador.Fields(2)
            End If
            rsAvaliador.MoveNext
        Loop
        PontosColabFor = pontoMaior
        If PesoColab > PesoMatriz Then PontosColabFor = 100
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
    If CONTADOR > 0 Then
        If ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor + PontosColabADP) / CONTADOR) >= MediaGlobal Then
            chamaForm.Label41.ForeColor = &H8000&
        ElseIf ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor) + PontosColabADP / CONTADOR) < MediaGlobal And ((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor) + PontosColabADP / CONTADOR) >= vAprovadoRest Then
            chamaForm.Label41.ForeColor = &H80FF&
        Else
            chamaForm.Label41.ForeColor = &HC0&
        End If
        
        If CONTADOR > 0 Then
            chamaForm.Label41 = Format(((PontosColabExp + PontosTotaisHab + PontosTotaisTrei + PontosColabFor + PontosColabADP) / CONTADOR), "#,##0.00;(#,##0.00)") & " %"
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
    
    If rsTmpExp.RecordCount <> 0 Then frmMenu2.ProgressBar1.Max = rsTmpExp.RecordCount
    X = 0
    Legenda = "Aguarde, reavaliando experiência dos colaboradores..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    'While Not rsTmpExp.EOF
    'Do While Not rsTmpExp.EOF
    For Y = 1 To rsTmpExp.RecordCount
        frmMenu2.ProgressBar1.Value = X
        periodoEmMeses = DateDiff("m", rsTmpExp.Fields(5), Now)
        If Val(periodoEmMeses) > 0 Then
            registraExperiencia rsTmpExp.Fields(0), rsTmpExp.Fields(3), periodoEmMeses
        End If
        rsTmpExp.MoveNext
        X = X + 1
    Next
    'Loop
    'Wend
    frmMenu2.ProgressBar1.Value = 0
    frmMenu2.StatusBar1.Panels(3).Text = "Grupo: " & GrupoUsu
    Set rsVExp = Nothing
    rsTmpExp.Close
    Set rsTmpExp = Nothing
End Function

Public Sub afastaColaborador()
    Dim rsAfastColab As New ADODB.Recordset
    Dim SqlAfastColab As String
    'RETORNA Colaborador
    If MeuLV.Label2 = "Afastados" Then
        frmRetorna.Show 1
        
'        Dim dataret As String
'        dataret = InputBox("Digite a data de retorno do Colaborador: Use o formato: DD/MM/AAAA", "SGC", Date)
'        If dataret <> "" Then
'            If IsDate(dataret) Then
'                SqlAfastColab = "Update tbcolaboradores set Ativo = CASE WHEN Ativo = 'S' then 'A' WHEN Ativo = 'A' then 'S' END where cpf = '" & varGlobal & "';" & _
'                "Update tbcolaboradores set dataafastamento = null where cpf = '" & varGlobal & "'"
'                rsAfastColab.Open SqlAfastColab, cnBanco
'                Msgbox "Colaborador retornou do afastamento"
'            Else
'                Msgbox " Data Inválida ", vbCritical, "SGC"
'            End If
'        End If
    'AFASTA Colaborador
    Else
        Dim data As String
        data = InputBox("Digite a data de afastamento do Colaborador: Use o formato: DD/MM/AAAA", "SGC", Date)
        If data <> "" Then
            If IsDate(data) Then
                SqlAfastColab = "Update tbcolaboradores set dataafastamento = '" & Format(CStr(data), "YYYY/MM/DD") & "',Ativo = CASE WHEN Ativo = 'S' then 'A' WHEN Ativo = 'A' then 'S' END where cpf = '" & varGlobal & "'"
                rsAfastColab.Open SqlAfastColab, cnBanco
                MsgBox "Colaborador foi afastado"
            Else
                MsgBox " Data Inválida ", vbCritical, "SGC"
            End If
        End If
        
    End If
End Sub

Private Sub registraExperiencia(vCPF As String, vCodCargo As String, vTempoExp As Single)
    Dim rsExperiencia As New ADODB.Recordset
    Dim SqlExperiencia As String
    SqlVExp = "Select * from tbColaboradoresExp where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "' and nomeempresa = '" & NomeEmpresa & "' and codcargo = '" & vCodCargo & "'"
    rsVExp.Open SqlVExp, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsVExp.RecordCount > 0 Then
        SqlExperiencia = "Update tbColaboradoresExp set tempoexp = '" & Format(vTempoExp, "000") & " Meses" & "' Where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "' and nomeempresa = '" & NomeEmpresa & "' and codcargo = '" & vCodCargo & "'"
        rsExperiencia.Open SqlExperiencia, cnBanco
    Else
        SqlExperiencia = "Insert into tbColaboradoresExp(cpf,tipo,nomeempresa,codcargo,tempoexp,codcoligada) Values('" & vCPF & "','colaborador','" & NomeEmpresa & "','" & vCodCargo & "','" & Format(vTempoExp, "000") & " Meses" & "','" & vCodcoligada & "')"
        rsExperiencia.Open SqlExperiencia, cnBanco
    End If
    rsVExp.Close
End Sub

Public Sub CompoeComboTotvs(Combo As ComboBox, Tabela, campo, Campo1)
    On Error Resume Next
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
    sql = "Select * from " & Tabela & " Order By " & campo
    rsTabela.Open sql, cnBancoTotvs, adOpenKeyset, adLockOptimistic
    If Not rsTabela.EOF() Then
        Combo.Clear
        rsTabela.MoveFirst
        For X = 0 To rsTabela.RecordCount - 1
            If Not IsNull(rsTabela.Fields(Campo1)) Then Combo.AddItem rsTabela.Fields(Campo1)
            If Not IsNull(rsTabela.Fields(Campo1)) Then Combo.ItemData(Combo.NewIndex) = Val(rsTabela.Fields(0))
            rsTabela.MoveNext
        Next
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub

Public Sub CarregaComboTotvs(Tabela As String, campo As String, Filtro As String, Resultado As String, indice As Integer, Campo1 As String)
    On Error Resume Next
    Dim rsConsulta As New ADODB.Recordset
    Dim SqlConsulta As String
    SqlConsulta = "Select " & campo & "," & Campo1 & " from " & Tabela & " where " & campo & " = '" & Filtro & "' order by " & campo
    rsConsulta.Open SqlConsulta, cnBancoTotvs, adOpenKeyset, adLockOptimistic
    If Not rsConsulta.EOF Then rsConsulta.MoveFirst
    If rsConsulta.EOF Then
        MsgBox "Consulta não encontrada", vbInformation, "Totvs RM"
        chamaForm.Combo(indice + 1) = ""
        chamaForm.txtCons(indice) = ""
    Else
        chamaForm.txtCons(indice) = rsConsulta.Fields(0)
        chamaForm.Combo(indice + 1) = rsConsulta.Fields(1)
    End If
    rsConsulta.Close
    Set rsConsulta = Nothing
End Sub

Public Sub AchaComboTotvs(Combo As ComboBox, Tabela As String, campo As String, indice As Integer, Campo1 As String)
    Dim rsConsulta As New ADODB.Recordset
    Dim SqlConsulta As String
    SqlConsulta = "Select " & campo & " from " & Tabela & " where " & Campo1 & " = '" & Combo & "' order by " & Campo1
    rsConsulta.Open SqlConsulta, cnBancoTotvs, adOpenKeyset, adLockOptimistic
    If Not rsConsulta.EOF Then rsConsulta.MoveFirst
    chamaForm.txtCons(indice - 1) = rsConsulta.Fields(0)
    rsConsulta.Close
    Set rsConsulta = Nothing
End Sub

Public Sub GravaDadosDBTotvs(vRegistro As String)
On Error GoTo Err
    Dim rsGravaTotvs As New ADODB.Recordset
    Dim SqlGravaTotvs  As String
    Dim vCodpessoa As Integer, vCodImagem As Integer
    
    ConexaoTotvs
    
    'Inicia transação de dados
    cnBancoTotvs.BeginTrans
    
    'Verifica se há registro do colaborador no banco Totvs
    SqlGravaTotvs = "Select a.chapa,a.codfuncao,a.codsecao from pfunc as a where a.chapa = " & vRegistro
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    If rsGravaTotvs.RecordCount > 0 Then
        rsGravaTotvs.Close
        Set rsGravaTotvs = Nothing
        
        'Se houver registro cadastrado no banco Totvs, Apenas atualiza
        SqlGravaTotvs = "Update pfunc set codsecao = '" & vDadosTotvs(14) & "', codfuncao = '" & vDadosTotvs(13) & "' Where chapa = '" & vRegistro & "'"
        rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
        
        cnBancoTotvs.CommitTrans
        cnBancoTotvs.Close
        Set cnBancoTotvs = Nothing
        Exit Sub
    Else
        rsGravaTotvs.Close
        Set rsGravaTotvs = Nothing
    End If
    
    'Grava em PFUNC
    SqlGravaTotvs = "Select MAX(codigo)+1 from PPESSOA"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    vCodpessoa = rsGravaTotvs.Fields(0)
    rsGravaTotvs.Close
    Set rsGravaTotvs = Nothing
    
    'Grava em PPESSOA
    SqlGravaTotvs = "Insert into ppessoa(aluno,professor,usuariobiblios,funcionario,exfuncionario,candidato,codigo,nome,dtnascimento,sexo,grauinstrucao,carteiratrab) Values(0,0,0,0,0,0,'" & vCodpessoa & "','" & vDadosTotvs(1) & "','" & Format(CStr(vDadosTotvs(2)), "YYYY/MM/DD") & "','" & vDadosTotvs(5) & "','" & vDadosTotvs(6) & "','" & vDadosTotvs(3) & "')"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em PFUNC
    SqlGravaTotvs = "Insert into pfunc(codcoligada,codfilial,codpessoa,chapa,tipoadmissao,dataadmissao,motivoadmissao,codsindicato,codfuncao,codsecao,situacaorais,contribsindical,codrecebimento,codsituacao,codtipo,codhorario,jornadamensal,situacaofgts) Values(1,1,'" & vCodpessoa & "','" & vDadosTotvs(0) & "','" & vDadosTotvs(7) & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','" & vDadosTotvs(8) & "','" & vDadosTotvs(17) & "','" & vDadosTotvs(13) & "','" & vDadosTotvs(14) & "','" & vDadosTotvs(16) & "','" & vDadosTotvs(15) & "','" & vDadosTotvs(9) & "','" & vDadosTotvs(10) & "','" & vDadosTotvs(11) & "','" & vDadosTotvs(12) & "',2640,1)"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em PFHSTFCO
    SqlGravaTotvs = "Insert into PFHSTFCO(codcoligada,chapa,dtmudanca,motivo,codfuncao) Values(1,'" & vDadosTotvs(0) & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','01','" & vDadosTotvs(13) & "')"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em PFHSTSEC
    SqlGravaTotvs = "Insert into PFHSTSEC(codcoligada,chapa,dtmudanca,motivo,codsecao) Values(1,'" & vDadosTotvs(0) & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','01','" & vDadosTotvs(14) & "')"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em PFHSTSIT
    SqlGravaTotvs = "Insert into PFHSTSIT(codcoligada,chapa,datamudanca,motivo,novasituacao) Values(1,'" & vDadosTotvs(0) & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','" & vDadosTotvs(8) & "','A')"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em PFHSTSAL
    SqlGravaTotvs = "Insert into PFHSTSAL(codcoligada,chapa,dtmudanca,datadereferencia,motivo,nrosalario,salario,jornada) Values(1,'" & vDadosTotvs(0) & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','01',1,44.00,2640)"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em PFHSTHOR
    SqlGravaTotvs = "Insert into PFHSTHOR(codcoligada,chapa,dtmudanca,codhorario,indiniciohor,comportamentohorarioanterior,comportamentohorarioatual) Values(1,'" & vDadosTotvs(0) & "','" & Format(CStr(Date), "YYYY/MM/DD") & "','0003',6,0,0)"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    'Grava em GIMAGEM
    
    SqlGravaTotvs = "Select MAX(ID)+1 from GIMAGEM"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    vCodImagem = rsGravaTotvs.Fields(0)
    rsGravaTotvs.Close
    Set rsGravaTotvs = Nothing
    
    
    Set mStream = New ADODB.Stream
    mStream.Type = adTypeBinary
    mStream.Open
    mStream.LoadFromFile vDadosTotvs(4)
    
    SqlGravaTotvs = "Select * from GIMAGEM"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs, adOpenKeyset, adLockOptimistic
    rsGravaTotvs.AddNew
    rsGravaTotvs.Fields(0) = vCodImagem
    rsGravaTotvs.Fields(1) = "P"
    rsGravaTotvs.Fields(2) = mStream.Read
    rsGravaTotvs.Update
    rsGravaTotvs.Close
    Set rsGravaTotvs = Nothing
    
    SqlGravaTotvs = "Update PPESSOA set idimagem = '" & vCodImagem & "',AjustaTamanhoFoto = 1 Where codigo = '" & vCodpessoa & "'"
    rsGravaTotvs.Open SqlGravaTotvs, cnBancoTotvs
    
    cnBancoTotvs.CommitTrans
    cnBancoTotvs.Close
    Set cnBancoTotvs = Nothing
    Exit Sub
Err:
    MsgBox "A gravação dos dados totvs não foi realizada com sucesso", vbCritical, "SGCH"
    cnBancoTotvs.RollbackTrans
    Exit Sub
End Sub


Public Sub ateraCargoTotvs(vRegistro As String, vFuncao As String)
    ConexaoTotvs
    'Inicia transação de dados
    cnBancoTotvs.BeginTrans
    
    Dim rsAlteraColSGCH As New ADODB.Recordset
    Dim SqlAlteraColSGCH As String

    SqlAlteraColSGCH = "Update pfunc set codsecao = '001.01.01.02', codfuncao = '" & vFuncao & "' Where chapa = '" & vRegistro & "'"
    rsAlteraColSGCH.Open SqlAlteraColSGCH, cnBancoTotvs
    
    'Finaliza transação de dados
    cnBancoTotvs.CommitTrans
    cnBancoTotvs.Close
    Set cnBancoTotvs = Nothing
    
End Sub

Public Sub buscaDemitidos()
    ConexaoTotvs
    'Inicia transação de dados
    cnBancoTotvs.BeginTrans
    
    Dim rsbuscaDemitidos As New ADODB.Recordset
    Dim SqlbuscaDemitidos As String

    Dim rsGravaDemitidos As New ADODB.Recordset
    Dim SqlGravaDemitidos As String

    SqlbuscaDemitidos = "select * from zSGCH_Demitidos where controleSGCH = ''"
    rsbuscaDemitidos.Open SqlbuscaDemitidos, cnBancoTotvs, adOpenKeyset, adLockReadOnly
    If Not rsbuscaDemitidos.EOF Then rsbuscaDemitidos.MoveFirst
    While Not rsbuscaDemitidos.EOF
        SqlGravaDemitidos = "Update tbColaboradores set ativo = 'N',datarecisao = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))),homologacaonum = 'Ver Totvs',homologacaoorgao = 'Ver Totvs' Where codcoligada = '" & vCodcoligada & "' and codcolaborador = '" & rsbuscaDemitidos.Fields(0) & "'"
        rsGravaDemitidos.Open SqlGravaDemitidos, cnBanco
        rsbuscaDemitidos.MoveNext
    Wend
    rsbuscaDemitidos.Close
    Set rsbuscaDemitidos = Nothing

    SqlbuscaDemitidos = "Update zSGCH_Demitidos set controleSGCH = 'V'"
    rsbuscaDemitidos.Open SqlbuscaDemitidos, cnBancoTotvs
    
    'Finaliza transação de dados
    cnBancoTotvs.CommitTrans
    cnBancoTotvs.Close
    Set cnBancoTotvs = Nothing
End Sub

Public Sub ajusta_LV()
On Error Resume Next
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
    cnBanco.BeginTrans
   
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
End Sub

Public Function MsgComboBox(Texto As String, Optional Botoes As Integer = 0, Optional Letra As Integer = 10, Optional Cor As Variant = vbBlack, Optional Foco As Integer = 0) As Boolean
    'If bMensagemAberta = True Then Exit Function
    'bMensagemAberta = True
    MsgComboBox = False
    With frmMsgComboBox
        .txtTexto = Texto
        .txtTexto.FontSize = Letra
        .txtTexto.ForeColor = Cor
        Select Case Botoes
            Case 0
                .Combo1.Clear
                .Combo1.AddItem "Todos"
                .Combo1.AddItem "Aberto"
                .Combo1.AddItem "Fechado"
                '.Combo1.AddItem "Em andamento"
                .Combo1.AddItem "Cancelada"
                '.cmdopcao1.Caption = "&OK"
                '.cmdopcao2.Visible = False
            'Case 1
            '    .cmdopcao1.Caption = "&Sim"
            '    .cmdopcao2.Caption = "&Não"
            'Case 2
            '    .cmdopcao1.Caption = "&Confirma"
            '    .cmdopcao2.Caption = "C&ancela"
        End Select
        If Foco <> 0 Then .cmdopcao1.TabIndex = 0
        .Show vbModal
        If TiPo = 1 Then
            MsgComboBox = True
        Else
            MsgComboBox = False
        End If
        
        
        'If .Retorno_MsgComboBox = True Then MsgComboBox = True
        'bMensagemAberta = False
    End With
End Function

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
                        'LV.ListItems.Item(Y).Selected = True
                        'If vRaptor(Z) <> "" Then LV.SelectedItem.ListSubItems.Item(Z - 1) = vRaptor(Z)
                    End If
                Next
                Y = LV.ListItems.Count
                IncluirLV = True
'                Exit Function
            End If
        Next
        If Formulario = "Métodos & Processos" And LV.Name = "ListView1" Then
            'If separaDesLv(chamaForm.Text1.Text) = False Then
            '    IncluirLV = True
            '    Exit Function
            'Else
            '    IncluirLV = True
            'End If
        End If
        Set ItemLst = LV.ListItems.Add(, , vRaptor(1))
        Y = LV.ListItems.Count
    Else
        If Formulario = "Métodos & Processos" And LV.Name = "ListView1" Then
            'If separaDesLv(chamaForm.Text1.Text) = False Then
            '    IncluirLV = False
            '    Exit Function
            'Else
            '    IncluirLV = True
            'End If
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

Public Function chamaSQL(vSql As String)
    Sqlp = ""
    Sqlp = vSql
End Function


Public Function Compoe_Listview(LV As ListView, vSqlCompoe As String, vZerosEsq As String)
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
                'If chamaForm.Name = "frmMPCompleto" And LV.Name = "ListView1" And X = 1 Then
                '    If rsCompoe.Fields(X) = 0 Then
                '        ItemLst.SubItems(X) = "0"
                '    Else
                '        ItemLst.SubItems(X) = rsCompoe.Fields(X) '"" & Format(rsCompoe.Fields(X), "000000000")
                '    End If
                'Else
                    If rsCompoe.Fields(X) = 0 Then
                        ItemLst.SubItems(X) = "0"
                    Else
                        ItemLst.SubItems(X) = "" & rsCompoe.Fields(X)
                    End If
                'End If
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

'Função nativa do ZEUS, nao serve para o SGCH
'Essa rotina serve para verificar se o item/c.custo que esta sendo inserido no ListView1
'Ja está em uma outra OS
'Private Function separaDesLv(vTxtForm As String)
'    separaDesLv = True
'    Dim rsTransf As New ADODB.Recordset
'    Dim SqlTransf As String
'    Dim vCodLM As String, vCodSeq As String
'    Dim RECEBE As String
'    Dim CONTADOR As Integer, X As Integer
'    CONTADOR = 0
'    For X = 1 To Len(vTxtForm)
'        If Mid(vTxtForm, X, 1) = ";" Then
'            If Len(RECEBE) = 5 Then
'                vCodLM = Mid$(RECEBE, 1, 2)
'                vCodSeq = Mid$(RECEBE, 3, 3)
'            Else
'                vCodLM = Mid$(RECEBE, 1, 2)
'                vCodSeq = Mid$(RECEBE, 3, 4)
'            End If
''            vCodLM = Mid$(RECEBE, 1, 2)
''            vCodSeq = Mid$(RECEBE, 3, 3)
'            SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
'            "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoTotvs & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
'            "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoTotvs & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
'            "inner join tbProjetos as g on g.codprojeto = d.codprojeto where a.fce = '" & Val(chamaForm.txtformula(12)) & "' and a.projeto = '" & chamaForm.txtformula(13).Text & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "' and a.idcc = '" & chamaForm.txtformula(0) & "' and a.idoperacao ='" & chamaForm.Combo1 & "'"
'            rsTransf.Open SqlTransf, cnBanco, adOpenKeyset, adLockReadOnly
'            If rsTransf.RecordCount > 0 Then
'                mobjMsg.Abrir "Desenho: " & rsTransf.Fields(8) & vbCrLf & _
'                              "Posição: " & rsTransf.Fields(11) & vbCrLf & _
'                              "Item:" & rsTransf.Fields(12) & vbCrLf & _
'                              "C.Custo:" & rsTransf.Fields(6) & vbCrLf & _
'                              "Registrado na OS:" & Format(rsTransf.Fields(0), "000000000") & " - Programação: " & Format(rsTransf.Fields(7), "000000"), Ok, critico, "Atenção"
'                separaDesLv = False
'                rsTransf.Close
'                Set rsTransf = Nothing
'                Exit Function
'            End If
'            rsTransf.Close
'            Set rsTransf = Nothing
'            RECEBE = ""
'        Else
'            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
'        End If
'    Next
'    If RECEBE <> "" Then
'        If Len(RECEBE) = 5 Then
'            vCodLM = Mid$(RECEBE, 1, 2)
'            vCodSeq = Mid$(RECEBE, 3, 3)
'        Else
'            vCodLM = Mid$(RECEBE, 1, 2)
'            vCodSeq = Mid$(RECEBE, 3, 4)
'        End If
''        vCodLM = Mid$(RECEBE, 1, 2)
''        vCodSeq = Mid$(RECEBE, 3, 3)
'        SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
'        "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoTotvs & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
'        "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoTotvs & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
'        "inner join tbProjetos as g on g.codprojeto = d.codprojeto where a.fce = '" & Val(chamaForm.txtformula(12)) & "' and a.projeto = '" & chamaForm.txtformula(13).Text & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "' and a.idcc = '" & chamaForm.txtformula(0) & "' and a.idoperacao ='" & chamaForm.Combo1 & "'"
'        rsTransf.Open SqlTransf, cnBanco, adOpenKeyset, adLockReadOnly
'        If rsTransf.RecordCount > 0 Then
'            mobjMsg.Abrir "Desenho: " & rsTransf.Fields(8) & vbCrLf & _
'                          "Posição: " & rsTransf.Fields(11) & vbCrLf & _
'                          "Item:" & rsTransf.Fields(12) & vbCrLf & _
'                          "C.Custo:" & rsTransf.Fields(6) & vbCrLf & _
'                          "Registrado na OS:" & Format(rsTransf.Fields(0), "000000000") & " - Programação: " & Format(rsTransf.Fields(7), "000000"), Ok, critico, "Atenção"
'            separaDesLv = False
'            rsTransf.Close
'            Set rsTransf = Nothing
'            Exit Function
'        End If
'        rsTransf.Close
'        Set rsTransf = Nothing
'    End If
'End Function

'A Função abaixo é referente a ALTERAÇÃO de dados de qualquer ListView com até 15 colunas
Public Function AlteraLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
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

