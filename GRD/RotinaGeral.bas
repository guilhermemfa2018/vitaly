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
Public vStatusMedicao As String

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
Public vBanco As String  'Armazena nome do banco SAP
Public vUsuBancoTovs As String  'Armazena usuario do banco SAP
Public vSenhaBancoSAP As String  'Armazena senha do banco SAP

Public chamaForm As Form

'Public MeuLV As New frmPesqGeral
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
Public vMotivoBloqueio As Boolean
Public vTextMotivoBloqueio As String
Public vValorAcumulado As Double

Public Sub Main()
On Error GoTo Err1
    'frmSplash.Show
    'Conexao
    'MDIPrincipal.Show
    Exit Sub
Err1:
    MsgBox "(ADO) Erro ao tentar acessar DB " & vbCrLf & Err.Number & " - Procure o administrador da rede " & Err.Description, 16, "Mensagem de erro"
    'mobjMsg.Abrir "(ADO) Erro ao tentar acessar DB " & vbCrLf & Err.Number & " - Procure o administrador da rede " & Err.Description, 16, "Mensagem de erro", ok, critico, "Atenção"
    Exit Sub
End Sub

Public Function Conexao()
On Error GoTo Err1
    Conexao = True
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & frmListaGRD.Text4.Text & ";Persist Security Info=True;Connect Timeout=0;User ID=" & frmListaGRD.Text5.Text & ";Initial Catalog=" & frmListaGRD.Text6.Text & ";Data Source=" & frmListaGRD.Text7.Text
    frmListaGRD.Label2.ForeColor = &H8000&
    frmListaGRD.Label2.Caption = "Conexão estabelecida com o servidor 10.10.10.232\srvweb"
    vBanco = frmListaGRD.Text6.Text
    Exit Function
Err1:
    frmListaGRD.Label2.ForeColor = &HFF&
    frmListaGRD.Label2.Caption = "Falha na conexão com o servidor 10.10.10.232\srvweb: " & Err.Number & " - " & Err.Description
    Conexao = False
    Exit Function
End Function

'ABAIXO CONEXÃO COM O BANCO DE DADOS RM
'Public Function ConexaoSAP()
'On Error GoTo Err1
'    Set cnBancoSAP = New ADODB.Connection
'    cnBancoSAP.Open "Provider=SQLOLEDB.1;Password=" & vSenhaBancoSAP & ";Persist Security Info=True;User ID=" & vUsuBancoTovs & ";Initial Catalog=" & vBancoSAP & ";Data Source=" & vServerSAP
'    vIntegra = "S"
'    'achaSecaoZEUSH
'    'criaTrigger
'    Exit Function
'Err1:
'    MsgBox "Erro de conexão com Banco SAP", vbCritical, "Atenção"
'    'mobjMsg.Abrir "Erro de conexão com Banco SAP", Ok, critico, "Atenção"
'    vIntegra = "N"
'    Exit Function
'End Function


Public Function ConexaoLdap()
    Dim sUser As String, sDN As String, sRoot As String
    sUser = "admin"
    sDN = "uid=" & sUser & ",ou=usuarios,dc=id"
    sRoot = "LDAP://10.10.10.29/phpldapadmin:389"
    Dim oDS: Set oDS = GetObject("LDAP:")
    'On Error GoTo AuthError
'    Dim oAuth: Set oAuth = oDS.OpenDSObject(sRoot, sDN, "049332id", &H200)
    Dim oAuth: Set oAuth = oDS.OpenDSObject("LDAP://10.10.10.29:389/ou=usuarios,dc=id", "uid=admin", "049332id", 1)
    'On Error GoTo 0
    MsgBox "Login Successful"
    Exit Function
AuthError:
    If Err.Number = -2147023570 Then
        MsgBox "Wrong Username or password !!!"
    End If
    On Error GoTo 0
End Function


'Public Sub CompoeCombo(Combo As ComboBox, Tabela, Campo1)
'    Dim sql As String
'    Dim rsTabela As New ADODB.Recordset
'    Dim X As Integer
'    'Se a tabela for ID_APROP_PERIODO, somente irá exibir os 3 últimos períodos
'    If Tabela = "ID_APROP_PERIODO" Then
'        sql = "Select top 10 CONVERT(VARCHAR,a.DTINICIAL,103) + ' a ' + CONVERT(VARCHAR,a.DTFINAL,103) AS PERIODO from " & vBancoSAP & ".DBO." & Tabela & " as a Order By a." & Campo1 & " desc"
'    ElseIf Tabela = "GFILIAL" Then
'        sql = "Select CAST(A.CODFILIAL AS VARCHAR) + ' - ' + A.NOMEFANTASIA AS FILIAL from " & vBancoSAP & ".DBO." & Tabela & " as a where codcoligada = 1 Order By a." & Campo1 & " ASC"
'    Else
'        sql = "Select * from " & Tabela & " where codcoligada = '" & vCodcoligada & "' Order By " & Campo1
'    End If
'    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
'    If Not rsTabela.EOF() Then
'        Combo.Clear
'        rsTabela.MoveFirst
'        For X = 0 To rsTabela.RecordCount - 1
'            Combo.AddItem rsTabela.Fields(0)
'            rsTabela.MoveNext
'        Next
'    End If
'    Combo.ItemData(Combo.NewIndex) = 1
'    rsTabela.Close
'    Set rsTabela = Nothing
'End Sub

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
            Combo.ItemData(Combo.NewIndex) = val(rsTabela.Fields(0))
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
            Combo.ItemData(Combo.NewIndex) = val(rsTabela.Fields(0))
            rsTabela.MoveNext
        Next
    End If
    rsTabela.Close
    Set rsTabela = Nothing
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

Public Sub CompoeCampoCombo(Codigo, Combo As ComboBox)
    Dim X As Integer
    For X = 0 To Combo.ListCount - 1
        Combo.ListIndex = X
        If Combo.List(X) = Codigo Then
            Exit For
        End If
    Next
End Sub

'Public Sub CompoeComboLV(Combo As ComboBox, Optional Column As ColumnHeader = Nothing)
'    Dim c As ColumnHeader
'    Combo.Clear
'    'If Column Is Nothing Then
'        For Each c In MeuLV.ListView1.ColumnHeaders
'            Combo.AddItem c
'        Next
'        Combo.Text = Combo.List(0)
'    'End If
'End Sub

'Public Sub CompoeComboLVPesq(Combo As ComboBox, LV As ListView, vIndiceCombo As Integer, Optional Column As ColumnHeader = Nothing)
'    Dim c As ColumnHeader
'    'If Column Is Nothing Then
'        For Each c In LV.ColumnHeaders
'            Combo.AddItem c
'        Next
'        Combo.Text = Combo.List(vIndiceCombo)
'    'End If
'End Sub

'Public Sub CompoeComboCC(Combo As ComboBox)
'    Dim sql As String
'    Dim rsTabela As New ADODB.Recordset
'    Dim X As Integer
'    sql = "select a.NOME from CORPORERM.dbo.GCCUSTO as a where a.ATIVO = 'T' and substring(a.nome,1,4) = '3000' or substring(a.nome,1,4) = '4000' or substring(a.nome,1,4) = '7000' or substring(a.nome,1,4) = '5000'"
'    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
'    Combo.Clear
'    If Not rsTabela.EOF() Then
'        rsTabela.MoveFirst
'        For X = 0 To rsTabela.RecordCount - 1
'            Combo.AddItem rsTabela.Fields(0)
'            rsTabela.MoveNext
'        Next
'    Else
'        Combo.AddItem ("-")
'        Combo.Text = "-"
'    End If
'    rsTabela.Close
'    Set rsTabela = Nothing
'End Sub

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
        
        SqlAvaliador = "select * from tbMatrizExp as a left join tbColaboradoresExp as b on a.codcargo = b.codcargo and a.codmatriz = '" & val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and b.cpf = '" & chamaForm.mskCadMatriz & "' and     b.tipo = '" & Tipo & "' where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
        rsAvaliador.Open SqlAvaliador, cnBanco, adOpenKeyset, adLockOptimistic
        ContCargoMatExp = 0
        PontosMatrizExp = 0
        PontosColabExp = 0
        '>>Soma todos os pontos de EXPERIENCIA da matriz
        If rsAvaliador.RecordCount > 0 Then
            If Mid$(rsAvaliador.Fields(2), 4, 4) = "Anos" Then
                'Converte anos para meses
                ConverTido = val(Mid$(rsAvaliador.Fields(2), 1, 3)) * 12
            Else
                ConverTido = val(Mid$(rsAvaliador.Fields(2), 1, 3))
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
                    ConverTido = val(Mid$(rsAvaliador.Fields(8), 1, 3)) * 12
                Else
                    ConverTido = val(Mid$(rsAvaliador.Fields(8), 1, 3))
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
        SqlAvaliador = "select a.cpf,a.codmatriz,a.codhabilidade,a.pontuacao,b.peso from tbColaboradoresHab as a inner join tbHabilidades as b on a.codcoligada = '" & vCodcoligada & "' and a.codhabilidade = b.codhabilidade and a.codmatriz = '" & val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' and a.cpf = '" & chamaForm.mskCadMatriz & "' and     a.tipo = '" & Tipo & "' where a.codmatriz = '" & val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
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
        SqlAvaliador = "select * from tbMatrizCur as a left join tbcolaboradoresCur as b on a.codtreinamento = b.codtreinamento and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & Tipo & "' and b.codnivel >= a.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "'"
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
        SqlAvaliador = "select c.codmatriz,c.codescolaridade,c.pontuacao,b.cpf,b.tipo,b.codescolaridade,a.peso from tbescolaridade as a left join tbcolaboradoresesc as b on a.codescolaridade = b.codescolaridade and b.cpf = '" & chamaForm.mskCadMatriz & "' and b.tipo = '" & Tipo & "' left join tbmatrizEsc as c on a.codescolaridade = c.codescolaridade and c.codmatriz = '" & val(Mid$(chamaForm.txtCadMatriz(4), 1, 6)) & "' where a.codcoligada = '" & vCodcoligada & "'"
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

Public Function Enter(Key As Integer) As Integer
   If Key = 13 Then
       Enter = 0
   Else
       Enter = Key
   End If
End Function

'A Função abaixo gera código para qualquer Tabela
Public Function GeraCodigoTB(vTabela As String, vCampo As String, vCampo2 As String, vText As String)
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    If vCampo2 = "" Then
        SqlGera = "Select top 1 * from " & vTabela & " order by " & vCampo & " Desc"
    Else
        SqlGera = "Select top 1 * from " & vTabela & " where " & vCampo2 & "=" & val(vText) & " order by " & vCampo & " Desc"
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


'A Função abaixo é referente a INCLUSÃO de dados de qualquer ListView com até 10 colunas
Public Function IncluirLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
    On Error Resume Next
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, z As Integer
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
                For z = 1 To LV.ColumnHeaders.Count
                    If z = 1 Then
                        If vRaptor(z) <> "" Then vRaptor(z) = LV.ListItems.Item(X)
                    Else
                        If vRaptor(z) <> "" Then LV.SelectedItem.ListSubItems.Item(z - 1) = vRaptor(z)
                    End If
                Next
                Y = LV.ListItems.Count
                IncluirLV = True
                Exit Function
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
    For z = 2 To LV.ColumnHeaders.Count
        If vRaptor(z) <> "" Then ItemLst.SubItems(z - 1) = vRaptor(z)
    Next
    If vRaptor(2).Visible = True And vRaptor(2).Enabled = True Then
        vRaptor(2).SetFocus
    Else
        vRaptor(3).SetFocus
    End If
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
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = " & val(vVar1) & ""
        Else
            SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & val(vVar1) & "' and " & vCampo2 & " = '" & val(vVar2) & "'"
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
    For X = 1 To vQtdCampos - 1
        If vQualquerDado(X + 1, 2) = "S" Then
            rsSalvar.Fields(X) = vQualquerDado(X + 1, 1)
        ElseIf vQualquerDado(X + 1, 2) = "I" Then
            rsSalvar.Fields(X) = val(vQualquerDado(X + 1, 1))
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
    Dim X As Integer, Y As Integer, z As Integer
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
        For z = 0 To LV.ColumnHeaders.Count
            If IsNumeric(vMatrix(z + 1)) Then
                If vMatrix(z + 1) = "0" Then
                    vQualquerDado(X, z + 1) = LV.ListItems.Item(X)
                Else
                    'Se o valor da Listview for igual a "-" grava zero
                    If LV.SelectedItem.ListSubItems.Item(val(vMatrix(z + 1))) = "-" Then
                        vQualquerDado(X, z + 1) = 0
                    Else
                        If LV.SelectedItem.ListSubItems.Item(val(vMatrix(z + 1))) = "" Then
                            vQualquerDado(X, z + 1) = " "
                        Else
                            vQualquerDado(X, z + 1) = LV.SelectedItem.ListSubItems.Item(val(vMatrix(z + 1)))
                        End If
                        
                        'vQualquerDado(X, Z + 1) = vMatrix(Z + 1)
                    End If
                End If
            Else
                vQualquerDado(X, z + 1) = vMatrix(z + 1)
            End If
        Next
    Next
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
    'If apontaLV = 9 And LV.Name = "Listview1" Then
    '    LV.SortKey = 11
    'Else
        LV.SortKey = 0
    'End If
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
    'Adiciona processo ao item selecionado no Listview1
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
