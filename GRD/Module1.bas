Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public CaminhoSkin As String

Public CorFundo As Long '--------------

Public RefreshVendas As Boolean
Public RefreshOS As Boolean

Public AgendaAberta As String
Public vcaminhodosom As String  'caminho ao arquivo do som para o form de alerta/aviso
Public CodCompromisso As String  'Passa o código do compromisso ao clicar na mensagem
Public NumPopUp As Integer
Public NumAlturaPopUp As Integer
Public NumDistPopUp As Integer

Public vDataBase As String, vDataCalc As String
Public vPeriodo As Integer

'MsgBox
Public Onde As String
Public Onde1 As String
'Valor da resposta da msgbox
Public Tp As Integer
'Verificar se input tem valor de retorno ou não
Public Res As Boolean
'Valor da resposta da inputmsg
Public Inp As String

'Public mobjMsg As MsgBox

'Public Tema As Msgbox
'Public SkinAtual As Msgbox

Public ResX As Single
Public ResY As Single
Public OldX As Single
Public OldY As Single
Public resolucao As Boolean

'muda data e símbolo de R$
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SCURRENCY = 20
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

' muda resolução do vídeo
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function GetClipCursor Lib "user32.dll" (lprc As RECT) As Long

Private Declare Function EnumDisplaySettings Lib "user32" Alias _
"EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, _
lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias _
"ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Dim DevM As DEVMODE
Public Data30 As Date
Public vEmailAprovadores As String, vNumMedicao As String, vColaborador As String, vPeriodoMedicao As String, vCompetenciaMedicao As String, vMotivoMedicao As String
Public vSubstituto As String, vMantemExpressao As String, vTituloFiltro As String
Public vIdFiltro As Integer

Public Sub ChangeRes(iWidth As Single, iHeight As Single)
   Dim A As Boolean
   Dim I As Long
   Do
      A = EnumDisplaySettings(0&, I, DevM)
      I = I + 1
   Loop Until (A = False)

   Dim B As Long
   DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
   DevM.dmPelsWidth = iWidth
   DevM.dmPelsHeight = iHeight
   B = ChangeDisplaySettings(DevM, 0)
End Sub

Public Function Img()
CaMinho = Servidor & ":"
Set Principal.Image1.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")
End Function

Function AlwaysOnTop(FrmID As Form, ByVal OnTop As Boolean) As Boolean
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop = True Then
        AlwaysOnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        AlwaysOnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
End Function

Public Function criaTabTemp()
On Error Resume Next
    'Criando uma tabela para exibição de Notas de Fornecedores
    Dim rsTabTemp As New ADODB.Recordset
    Dim SqlTabTemp As String
    SqlTabTemp = "CREATE TABLE TempNotaFornec(idfornecedor NVARCHAR(15) NOT NULL,nomefornecedor VARCHAR(200) NOT NULL,nota FLOAT NOT NULL, classificacao CHAR(1) NOT NULL)"
    rsTabTemp.Open SqlTabTemp, cnBanco
End Function

Public Function insereDadosTemp()
    On Error GoTo Err
1000    Dim rsVerLinguagem As New ADODB.Recordset
1100    Dim SqlVerLinguagem As String
1200    SqlVerLinguagem = "SELECT @@language"
1300    rsVerLinguagem.Open SqlVerLinguagem, cnBanco, adOpenKeyset, adLockReadOnly
1400    If rsVerLinguagem.Fields(0) = "us_english" Then vFormatoDatetime = "yyyy-mm-dd" Else vFormatoDatetime = "dd-mm-yyyy"
1500    rsVerLinguagem.Close
1600    Set rsVerLinguagem = Nothing
    
'1700    Dim vDataBase As String, vDataCalc As String
    
'1800    Dim vPeriodo As Integer
1900    Dim vNotaFornec As Double
2000    Dim vClassifica As String
    
    
2100    Dim rsParametros As New ADODB.Recordset
2200    Dim sqlParametros As String
        
2300    Dim rsGravaFornec As New ADODB.Recordset
2400    Dim sqlGravaFornec As String
    
2500    Dim rsGeraDados As New ADODB.Recordset
2600    Dim sqlGeraDados As String
    
2700    Dim rsClassificacao As New ADODB.Recordset
2800    Dim sqlClassificacao As String
    
2900    Dim Y As Integer, X As Integer
    
3000    Dim rsDeletaTemp As New ADODB.Recordset
3100    Dim sqlDeletaTemp As String
    
    
3200    sqlParametros = "Select a.aprovadorest from tbparametros as a"
3300    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockReadOnly
3400    vPeriodo = rsParametros.Fields(0)
    
'3500    vDataBase = Date
'3600    vDataCalc = CDate(vDataBase) - (vPeriodo * 30)
    
3700    vPeridoAvFornec = vDataCalc
3800    rsParametros.Close
3900    Set rsParametros = Nothing
    
'4000    If frmPrintRels.Frame2.Visible = True Then
'4000    If frmPrintRels Is Nothing Then
'4000     If frmPrintRels.DTPicker1.Value <> "" Then
'4100        vDataCalc = frmPrintRels.DTPicker1.Value
'4200        vDataBase = frmPrintRels.DTPicker2.Value
'4300    End If
    
    
4400    sqlGeraDados = "select top 500 a.CardCode as ID_Fornecedor,a.CardName as nome_fornecedor,avg(b.notaOC) as notaOC,'-' as classificacao from " & vBancoSAP & ".DBO.OPOR as a LEFT JOIN tbOCStatus as b on a.DocNum = b.docnum inner join tbfornecedores as c on a.CardCode COLLATE SQL_Latin1_General_CP1_CI_AS = c.idfornecedor " & _
                   "where b.dataavoc between '" & Format(vDataCalc, vFormatoDatetime) & "' and '" & Format(vDataBase, vFormatoDatetime) & "' and a.CANCELED = 'N' and c.status = 'Credenciado' and c.ativo = 'S' and b.notaoc is not null and b.statusoc <> '7' group by a.CardCode,a.CardName order by a.CardName asc"
4500    rsGeraDados.Open sqlGeraDados, cnBanco, adOpenKeyset, adLockReadOnly
    

4600    sqlDeletaTemp = "delete from TempNotaFornec"
4700    rsDeletaTemp.Open sqlDeletaTemp, cnBanco
    
4800    While Not rsGeraDados.EOF
4900        sqlClassificacao = "select * from tbClassificacao as a where a.de <= " & Replace(rsGeraDados.Fields(2), ",", ".") & " and a.para >= " & Replace(rsGeraDados.Fields(2), ",", ".")
5000        rsClassificacao.Open sqlClassificacao, cnBanco, adOpenKeyset, adLockReadOnly
        
5100        If rsGeraDados.Fields(2) > 100 Then
5200            vNotaFornec = 100
5300            vClassifica = "A"
5400        Else
5500            vNotaFornec = rsGeraDados.Fields(2)
5600            vClassifica = rsClassificacao.Fields(0)
5700        End If
        
5800        sqlGravaFornec = "INSERT INTO TempNotaFornec(idfornecedor,nomefornecedor,nota,classificacao) VALUES('" & rsGeraDados.Fields(0) & "','" & rsGeraDados.Fields(1) & "'," & Replace(vNotaFornec, ",", ".") & ",'" & vClassifica & "')"
5900        rsGravaFornec.Open sqlGravaFornec, cnBanco
        
6000        rsClassificacao.Close
        
6100        rsGeraDados.MoveNext
6200    Wend
    

6300    rsGeraDados.Close
6400    Set rsGeraDados = Nothing
    Exit Function
Err:
    If Err.Number <> 3705 Then
        If Err.Number <> 727 And Err.Number <> 91 Then
            sqlGeraDados = "select top 500 a.CardCode as ID_Fornecedor,a.CardName as nome_fornecedor,avg(b.notaOC) as notaOC,'-' as classificacao from " & vBancoSAP & ".DBO.OPOR as a LEFT JOIN tbOCStatus as b on a.DocNum = b.docnum inner join tbfornecedores as c on a.CardCode COLLATE SQL_Latin1_General_CP1_CI_AS = c.idfornecedor " & _
                           "where b.dataavoc between '" & Format(frmPrintRels.DTPicker1.Value, vFormatoDatetime) & "' and '" & Format(frmPrintRels.DTPicker2.Value, vFormatoDatetime) & "' and a.CANCELED = 'N' and c.status = 'Credenciado' and c.ativo = 'S' and b.notaoc is not null group by a.CardCode,a.CardName order by a.CardName asc"
            rsGeraDados.Open sqlGeraDados, cnBanco, adOpenKeyset, adLockReadOnly
        Else
            MsgBox "grava_Dados Nº do erro: " & Err.Number & ", na linha: " & Str$(Erl) & vbCrLf & _
            " Descrição: " & Err.Description
        End If
    End If
    Resume Next
End Function

Public Function montaDadosClassifica()
    Dim rsClassificacao As New ADODB.Recordset
    Dim sqlClassificacao As String
    Dim X As Integer
    sqlClassificacao = "Select * from tbclassificacao as a order by idclassificacao"
    rsClassificacao.Open sqlClassificacao, cnBanco, adOpenKeyset, adLockReadOnly
    X = 0
    While Not rsClassificacao.EOF
        vQualquerDado(X, 1) = rsClassificacao.Fields(0)
        vQualquerDado(X, 2) = rsClassificacao.Fields(1)
        vQualquerDado(X, 3) = rsClassificacao.Fields(2)
        X = X + 1
        rsClassificacao.MoveNext
    Wend
End Function

Public Function CriarTabelasADOOFFLine(vServer As String, vBanco As String, vUsuario As String, vSenha As String, vQualBanco As String)
On Error GoTo Err1
    CriarTabelasADOOFFLine = False
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
    
    Set oConn = New ADODB.Connection
    
    
    oConn.Open "Provider=SQLOLEDB;Data Source=" & vServer & ";User ID=" & vUsuario & ";Password=" & vSenha & ";"
    
    If vQualBanco = "FERRAMENTARIA_OFF" Then
        'CRIA BANCO Ferramentaria
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbDadosBanco(" & _
        "NomeServidor VARCHAR(50) NULL," & _
        "NomeBanco VARCHAR(50) NULL)"
    
        'TABELAS Ferramentaria
        '============================
        'CRIA TABELAS Ferramentaria
    
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbConfiguracoes(" & _
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
        
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbMenuConf(" & _
        "idmenu NUMERIC NOT NULL," & _
        "idsub VARCHAR(10) NOT NULL," & _
        "tipo VARCHAR(20) NOT NULL," & _
        "nome VARCHAR(50) NOT NULL," & _
        "id INT NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "icon INT NOT NULL," & _
        "PRIMARY KEY (idmenu,idsub,tipo,codcoligada))"
        
    
        'CRIA TABELAS ESPECIFICAS DO SISTEMA
    
        'Tabela de emprestimo
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbEmprestimo(" & _
        "chapa VARCHAR(20) NOT NULL," & _
        "nome VARCHAR(100) NOT NULL," & _
        "codfuncao VARCHAR(10) NOT NULL," & _
        "nomefuncao VARCHAR(100) NOT NULL," & _
        "codsecao VARCHAR(35) NOT NULL," & _
        "nomesecao VARCHAR(100) NOT NULL," & _
        "dataemprestimo DATETIME NOT NULL," & _
        "idmov INT NOT NULL," & _
        "numeromov VARCHAR(35) NOT NULL," & _
        "serie VARCHAR(10) NOT NULL," & _
        "status VARCHAR(1) NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "localestoque VARCHAR(100) NOT NULL," & _
        "nomequememprestou VARCHAR(80) NOT NULL," & _
        "codusuariorm VARCHAR(50) NOT NULL," & _
        "PRIMARY KEY (chapa,numeromov,serie,codcoligada))"
    
        'Tabela de itens emprestados
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbEmprestimoItens(" & _
        "chapa VARCHAR(20) NOT NULL," & _
        "localestoque VARCHAR(100) NOT NULL," & _
        "codigoprd VARCHAR(100) NOT NULL," & _
        "descricao VARCHAR(300) NOT NULL," & _
        "idmov INT NOT NULL," & _
        "idprd VARCHAR(100) NOT NULL," & _
        "qtdemprestado NUMERIC NOT NULL," & _
        "qtddevolvida NUMERIC NOT NULL," & _
        "qtdpendente NUMERIC NOT NULL," & _
        "dataemprestimo DATETIME NOT NULL," & _
        "horaemprestimo DATETIME NOT NULL," & _
        "status VARCHAR(1) NOT NULL," & _
        "nomequememprestou VARCHAR(80) NOT NULL," & _
        "numerosequencial INT NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "um VARCHAR(10) NOT NULL," & _
        "valortotal FLOAT NULL," & _
        "numeromov VARCHAR(35) NOT NULL," & _
        "serie VARCHAR(10) NOT NULL," & _
        "PRIMARY KEY (chapa,numeromov,numerosequencial,codcoligada))"
    
        'Tabela que gera NUMEROMOV
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbMov(" & _
        "numeromov VARCHAR(35) NOT NULL," & _
        "serie VARCHAR(8) NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (numeromov,serie,codcoligada))"
    
    
        'Tabela que registra as devoluções do sistema
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbDevolucao(" & _
        "chapa VARCHAR(20) NOT NULL," & _
        "nome VARCHAR(100) NOT NULL," & _
        "codfuncao VARCHAR(10) NOT NULL," & _
        "nomefuncao VARCHAR(100) NOT NULL," & _
        "codsecao VARCHAR(35) NOT NULL," & _
        "nomesecao VARCHAR(100) NOT NULL," & _
        "idmov INT NOT NULL," & _
        "numeromov VARCHAR(35) NOT NULL," & _
        "serie VARCHAR(10) NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "localestoque VARCHAR(100) NOT NULL," & _
        "nomequemrecebeu VARCHAR(80) NOT NULL," & _
        "codusuariorm VARCHAR(50) NOT NULL," & _
        "PRIMARY KEY (chapa,numeromov,codcoligada))"
    
        'Tabela que registra os itens devolvidos no sistema
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbDevolucaoItens(" & _
        "numeromov VARCHAR(35) NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "numerosequencial INT NOT NULL," & _
        "localestoque VARCHAR(100) NOT NULL," & _
        "codigoprd VARCHAR(100) NOT NULL," & _
        "descricao VARCHAR(300) NOT NULL," & _
        "idmov INT NOT NULL," & _
        "idprd VARCHAR(100) NOT NULL," & _
        "qtddevolvida NUMERIC NOT NULL," & _
        "datadevolucao DATETIME NOT NULL," & _
        "horadevolucao DATETIME NOT NULL," & _
        "nomequememprestou VARCHAR(80) NOT NULL," & _
        "um VARCHAR(10) NOT NULL," & _
        "valortotal FLOAT NULL," & _
        "serie VARCHAR(10) NOT NULL," & _
        "idmovemp INT NOT NULL," & _
        "PRIMARY KEY (numeromov,numerosequencial,codcoligada,serie))"
    
        'Tabela que registra a sincronização das tabelas TMOV, TITMMOV, TMOVRELAC
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbSincronizacao(" & _
        "idmovsincronizado INT NOT NULL," & _
        "PRIMARY KEY (idmovsincronizado))"
        
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbTransfIDMov(" & _
        "idmov INT NOT NULL," & _
        "PRIMARY KEY (idmov))"
        
        'Tabela que registra os ultimos registros inseridos nas tabelas ON (registros a serem importados para as tabelas OFF)
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbControlInsertRec(" & _
        "nometabelaP VARCHAR(80) NOT NULL," & _
        "nometabelaS VARCHAR(80) NULL," & _
        "datacontrole1 DATETIME NOT NULL," & _
        "datacontrole2 DATETIME NULL," & _
        "PRIMARY KEY (nometabelaP))"
        
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbLocalEstoque(" & _
        "codigo NUMERIC NOT NULL," & _
        "codloc VARCHAR(15) NOT NULL," & _
        "nome VARCHAR(40) NULL," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (codigo,codloc))"
        
        '============================
        'TABELAS PADRAO
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbparametros(" & _
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
        
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbintegracao(" & _
        "tipobanco NUMERIC NOT NULL," & _
        "sistema NUMERIC NOT NULL," & _
        "modulo CHAR(10) NOT NULL," & _
        "nserver VARCHAR(50) NULL," & _
        "nbanco VARCHAR(50) NULL," & _
        "nusuario VARCHAR(50) NULL," & _
        "nsenha VARCHAR(50) NULL," & _
        "codcoligada INT NULL," & _
        "PRIMARY KEY (tipobanco,sistema,modulo))"
    
    
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbintegracaooffline(" & _
        "offline CHAR(1) NOT NULL," & _
        "nserver VARCHAR(50) NOT NULL," & _
        "nbanco VARCHAR(50) NOT NULL," & _
        "nusuario VARCHAR(50) NOT NULL," & _
        "nsenha VARCHAR(50) NOT NULL," & _
        "PRIMARY KEY (nserver))"
        
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbServerSincronizaDados(" & _
        "nserver VARCHAR(50) NOT NULL," & _
        "nbancototvs VARCHAR(50) NOT NULL," & _
        "nbancoferramentaria VARCHAR(50) NOT NULL," & _
        "nusuario VARCHAR(50) NOT NULL," & _
        "nsenha VARCHAR(50) NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (nserver,codcoligada))"
        
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbDadosEmpresa(" & _
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

        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbConfEmail(" & _
        "smtp VARCHAR(100) NULL," & _
        "usuario VARCHAR(50) NULL," & _
        "senha VARCHAR(30) NULL," & _
        "codcoligada INT NULL," & _
        "porta int NULL," & _
        "ssl int NULL," & _
        "smtpautentic int NULL)"

        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbLog(" & _
        "data VARCHAR(20) NULL," & _
        "hora VARCHAR(20) NOT NULL," & _
        "usuario VARCHAR(50) NOT NULL," & _
        "grupo VARCHAR(50) NULL," & _
        "formulario VARCHAR(50) NULL," & _
        "acao VARCHAR(300) NULL," & _
        "id INT NOT NULL IDENTITY," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (id,codcoligada))"
    
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbConfLV(" & _
        "nmusuario VARCHAR(50) NULL," & _
        "idmodulo NUMERIC NOT NULL," & _
        "indice NUMERIC NOT NULL," & _
        "posicao NUMERIC NOT NULL," & _
        "largura FLOAT NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "id INT NOT NULL IDENTITY," & _
        "PRIMARY KEY (id))"
    
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbConfGrupo(" & _
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
    
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbSenha(" & _
        "usuario VARCHAR(50) NOT NULL," & _
        "senha VARCHAR(50) NOT NULL," & _
        "codigo NUMERIC NOT NULL," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (codigo,codcoligada))"
    
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbMenu(" & _
        "idmenu NUMERIC NULL," & _
        "idsub VARCHAR(10) NULL," & _
        "tipo VARCHAR(20) NULL," & _
        "nome VARCHAR(50) NULL," & _
        "id INT NOT NULL IDENTITY," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (id,codcoligada))"
       
        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbUsuarios(" & _
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

        oConn.Execute "CREATE TABLE " & vBanco & ".dbo.tbGrupo(" & _
        "codigo NUMERIC NOT NULL," & _
        "descricao VARCHAR(50) NOT NULL," & _
        "ativo VARCHAR(1) NULL," & _
        "codcoligada INT NOT NULL," & _
        "PRIMARY KEY (codigo,codcoligada))"
    
        'ABAIXO: CRIA CONFIGURAÇÃO PARA USUÁRIO ADMINISTRADOR
        oConn.Close
    
        vCodcoligada = 1 'Primeiro cadastro de coligada
    
        oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenha & ";Persist Security Info=True;User ID=" & vUsuario & ";Initial Catalog=" & vBanco & ";Data Source=" & vServer

'        SqlSenha = "Insert into tbSenha(usuario,senha,codigo,codcoligada) Values('adm','123',1,'" & vCodcoligada & "');"
'        rsSenha.Open SqlSenha, oConn
    
'        SqlUsuario = "Insert into tbUsuarios(codigo,nome,codgrupo,ativo,codcoligada) Values(1,'Administrador do sistema',1,'S','" & vCodcoligada & "');"
'        rsUsuario.Open SqlUsuario, oConn
    
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
                       "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'6161','BUT','Sobre Ferramentaria','S','" & vCodcoligada & "',23);Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(1,7,'6162','BUT','Ajuda do IMRM','S','" & vCodcoligada & "',24);"
    
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
        CriarTabelasADOOFFLine = True
    Else
        
        'GAUTOINC
        oConn.Execute "use corporerm_OFF CREATE TABLE [" & vBanco & "].[dbo].[GAUTOINC]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL, " & _
        "   [CODSISTEMA] [varchar](2) NOT NULL, " & _
        "   [CODAUTOINC] [varchar](45) NOT NULL, " & _
        "   [VALAUTOINC] [int] NULL " & _
        ") ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [RECCREATEDBY] [varchar](50) NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [RECCREATEDON] [datetime] NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [RECMODIFIEDBY] [varchar](50) NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [RECMODIFIEDON] [datetime] NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [NOMETABELA] [varchar](50) NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [NOMECOLUNA] [varchar](50) NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD [IDPRDFCI] [int] NULL " & _
        "ALTER TABLE [dbo].[GAUTOINC] ADD  CONSTRAINT [PKGAUTOINC] " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CODSISTEMA] ASC,[CODAUTOINC] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"

        'TMOV
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TMOV]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[IDMOV] [int] NOT NULL,[CODFILIAL] [smallint] NULL,[CODLOC] [varchar](15) NULL,[CODLOCENTREGA] [varchar](15) NULL,[CODLOCDESTINO] [varchar](15) NULL,[CODCFO] [varchar](25) NULL,[CODCFONATUREZA] [varchar](25) NULL,[NUMEROMOV] [varchar](35) NULL,[SERIE] [varchar](8) NULL,[CODTMV] [varchar](10) NULL,[TIPO] [varchar](1) NULL,[STATUS] [varchar](1) NULL,[MOVIMPRESSO] [smallint] NULL,[DOCIMPRESSO] [smallint] NULL,[FATIMPRESSA] [smallint] NULL,[DATAEMISSAO] [datetime] NULL, " & _
        "   [DATASAIDA] [datetime] NULL,[DATAEXTRA1] [datetime] NULL,[DATAEXTRA2] [datetime] NULL,[CODRPR] [varchar](15) NULL,[COMISSAOREPRES] [FLOAT] NULL,[NORDEM] [varchar](20) NULL,[CODCPG] [varchar](5) NULL, [NUMEROTRIBUTOS] [smallint] NULL,[VALORBRUTO] [FLOAT] NULL,[VALORLIQUIDO] [FLOAT] NULL,[VALOROUTROS] [FLOAT] NULL,[OBSERVACAO] [varchar](60) NULL,[PERCENTUALFRETE] [FLOAT] NULL,[VALORFRETE] [FLOAT] NULL,[PERCENTUALSEGURO] [FLOAT] NULL,[VALORSEGURO] [FLOAT] NULL,[PERCENTUALDESC] [FLOAT] NULL, " & _
        "   [VALORDESC] [FLOAT] NULL,[PERCENTUALDESP] [FLOAT] NULL,[VALORDESP] [FLOAT] NULL,[PERCENTUALEXTRA1] [FLOAT] NULL,[VALOREXTRA1] [FLOAT] NULL,[PERCENTUALEXTRA2] [FLOAT] NULL,[VALOREXTRA2] [FLOAT] NULL,[PERCCOMISSAO] [FLOAT] NULL,[CODMEN] [varchar](5) NULL,[CODMEN2] [varchar](5) NULL,[VIADETRANSPORTE] [varchar](15) NULL,[PLACA] [varchar](10) NULL,[CODETDPLACA] [varchar](2) NULL,[PESOLIQUIDO] [FLOAT] NULL,[PESOBRUTO] [FLOAT] NULL,[MARCA] [varchar](10) NULL,[NUMERO] [int] NULL,[QUANTIDADE] [int] NULL, " & _
        "   [ESPECIE] [varchar](15) NULL,[CODTB1FAT] [varchar](10) NULL,[CODTB2FAT] [varchar](10) NULL,[CODTB3FAT] [varchar](10) NULL,[CODTB4FAT] [varchar](10) NULL,[CODTB5FAT] [varchar](10) NULL,[CODTB1FLX] [varchar](25) NULL,[CODTB2FLX] [varchar](25) NULL,[CODTB3FLX] [varchar](25) NULL,[CODTB4FLX] [varchar](25) NULL,[CODTB5FLX] [varchar](25) NULL,[IDMOVRELAC] [int] NULL,[IDMOVLCTFLUXUS] [int] NULL,[IDMOVPEDDESDOBRADO] [int] NULL,[CODMOEVALORLIQUIDO] [varchar](10) NULL,[DATABASEMOV] [datetime] NULL, " & _
        "   [DATAMOVIMENTO] [datetime] NULL,[NUMEROLCTGERADO] [smallint] NULL,[GEROUFATURA] [smallint] NULL,[NUMEROLCTABERTO] [smallint] NULL,[FLAGEXPORTACAO] [smallint] NULL,[EMITEBOLETA] [varchar](1) NULL,[CODMENDESCONTO] [varchar](5) NULL,[CODMENDESPESA] [varchar](5) NULL,[CODMENFRETE] [varchar](5) NULL,[FRETECIFOUFOB] [smallint] NULL,[USADESPFINANC] [smallint] NULL,[FLAGEXPORFISC] [smallint] NULL,[FLAGEXPORFAZENDA] [smallint] NULL,[VALORADIANTAMENTO] [FLOAT] NULL,[CODTRA] [varchar](5) NULL,[CODTRA2] [varchar](5) NULL, " & _
        "   [STATUSLIBERACAO] [smallint] NULL,[CODCFOAUX] [varchar](25) NULL,[IDLOT] [int] NULL,[ITENSAGRUPADOS] [smallint] NULL,[FLAGIMPRESSAOFAT] [varchar](1) NULL,[DATACANCELAMENTOMOV] [datetime] NULL,[VALORRECEBIDO] [FLOAT] NULL,[SEGUNDONUMERO] [varchar](20) NULL,[CODCCUSTO] [varchar](25) NULL,[CODCXA] [varchar](10) NULL,[CODVEN1] [varchar](16) NULL,[CODVEN2] [varchar](16) NULL,[CODVEN3] [varchar](16) NULL,[CODVEN4] [varchar](16) NULL,[PERCCOMISSAOVEN2] [FLOAT] NULL,[CODCOLCFO] [smallint] NULL, " & _
        "   [CODCOLCFONATUREZA] [smallint] NULL,[CODUSUARIO] [varchar](20) NULL,[CODFILIALENTREGA] [smallint] NULL,[CODFILIALDESTINO] [smallint] NULL,[FLAGAGRUPADOFLUXUS] [smallint] NULL,[CODCOLCXA] [INT] NULL,[GERADOPORLOTE] [smallint] NULL,[CODDEPARTAMENTO] [varchar](25) NULL,[CODCCUSTODESTINO] [varchar](25) NULL,[CODEVENTO] [smallint] NULL,[STATUSEXPORTCONT] [smallint] NULL,[CODLOTE] [int] NULL,[STATUSCHEQUE] [smallint] NULL,[DATAENTREGA] [datetime] NULL,[DATAPROGRAMACAO] [datetime] NULL, " & _
        "   [IDNAT] [int] NULL,[IDNAT2] [int] NULL,[CAMPOLIVRE1] [varchar](100) NULL,[CAMPOLIVRE2] [varchar](100) NULL,[CAMPOLIVRE3] [varchar](100) NULL,[GEROUCONTATRABALHO] [INT] NULL,[GERADOPORCONTATRABALHO] [INT] NULL,[HORULTIMAALTERACAO] [datetime] NULL,[CODLAF] [varchar](15) NULL,[DATAFECHAMENTO] [datetime] NULL,[NSEQDATAFECHAMENTO] [smallint] NULL,[NUMERORECIBO] [varchar](12) NULL,[IDLOTEPROCESSO] [int] NULL,[IDOBJOF] [varchar](20) NULL,[CODAGENDAMENTO] [int] NULL,[CHAPARESP] [varchar] (20) NULL, " & _
        "   [IDLOTEPROCESSOREFAT] [int] NULL,[INDUSOOBJ] [NUMERIC] (15,2) NULL,[SUBSERIE] [varchar](8) NULL,[STSCOMPRAS] [VARCHAR] (1) NULL,[CODLOCEXP] [INT] NULL,[IDCLASSMOV] [INT] NULL,[CODENTREGA] [INT] NULL,[CODFAIXAENTREGA] [INT] NULL,[DTHENTREGA] [DATETIME] NULL,[CONTABILIZADOPORTOTAL] [INT] NULL,[CODLAFE] [varchar](15) NULL,[IDPRJ] [int] NULL,[NUMEROCUPOM] [int] NULL,[NUMEROCAIXA] [int] NULL,[FLAGEFEITOSALDO] [smallint] NULL,[INTEGRADOBONUM] [INT] NULL,[CODMOELANCAMENTO] [varchar](10) NULL,[NAONUMERADO] [varchar](1) NULL, " & _
        "   [FLAGPROCESSADO] [INT] NULL,[ABATIMENTOICMS] [FLOAT] NULL,[TIPOCONSUMO] [smallint] NULL,[HORARIOEMISSAO] [datetime] NULL,[DATARETORNO] [datetime] NULL,[USUARIOCRIACAO] [varchar](20) NULL,[DATACRIACAO] [datetime] NULL,[IDCONTATOENTREGA] [int] NULL,[IDCONTATOCOBRANCA] [int] NULL,[STATUSSEPARACAO] [varchar](1) NULL,[STSEMAIL] [INT] NULL,[VALORFRETECTRC] [FLOAT] NULL,[PONTOVENDA] [varchar](10) NULL,[PRAZOENTREGA] [int] NULL,[VALORBRUTOINTERNO] [FLOAT] NULL,[IDAIDF] [smallint] NULL,[IDSALDOESTOQUE] [int] NULL, " & _
        "   [VINCULADOESTOQUEFL] [INT] NULL,[IDREDUCAOZ] [int] NULL,[HORASAIDA] [datetime] NULL,[CODMUNSERVICO] [varchar](20) NULL,[CODETDMUNSERV] [varchar](2) NULL,[APROPRIADO] [smallint] NULL,[CODIGOSERVICO] [varchar](15) NULL,[DATADEDUCAO] [datetime] NULL,[CODDIARIO] [varchar](5) NULL,[SEQDIARIO] [varchar](9) NULL,[SEQDIARIOESTORNO] [varchar](9) NULL,[INSSEMOUTRAEMPRESA] [FLOAT] NULL,[IDMOVCTRC] [int] NULL,[DATAPROGRAMACAOANT] [datetime] NULL,[CODTDO] [varchar](10) NULL,[VALORDESCCONDICIONAL] [FLOAT] NULL, " & _
        "   [VALORDESPCONDICIONAL] [FLOAT] NULL,[CODIGOIRRF] [varchar](10) NULL,[DEDUCAOIRRF] [FLOAT] NULL,[PERCENTBASEINSS] [FLOAT] NULL,[PERCBASEINSSEMPREGADO] [FLOAT] NULL,[CONTORCAMENTOANTIGO] [FLOAT] NULL,[CODDEPTODESTINO] [varchar](25) NULL,[DATACONTABILIZACAO] [datetime] NULL,[CODVIATRANSPORTE] [varchar](1) NULL,[VALORSERVICO] [FLOAT] NULL,[SEQUENCIALESTOQUE] [int] NULL,[DISTANCIA] [int] NULL,[UNCALCULO] [varchar](5) NULL,[FORMACALCULO] [varchar](1) NULL,[INTEGRADOAUTOMACAO] [smallint] NULL, " & _
        "   [INTEGRAAPLICACAO] [char](1) NOT NULL,[CLASSECONSUMO] [varchar](1) NULL,[TIPOASSINANTE] [varchar](1) NULL,[FASE] [varchar](1) NULL,[TIPOUTILIZACAO] [varchar](1) NULL,[GRUPOTENSAO] [varchar](1) NULL,[DATALANCAMENTO] [datetime] NULL,[EXTENPORANEO] [FLOAT] NULL,[RECIBONFESTATUS] [varchar](1) NULL,[RECIBONFETIPO] [smallint] NULL,[RECIBONFENUMERO] [varchar](12) NULL,[RECIBONFESITUACAO] [smallint] NULL,[IDMOVCFO] [int] NULL,[OCAUTONOMO] [smallint] NULL,[VALORMERCADORIAS] [FLOAT] NULL, " & _
        "   [NATUREZAVOLUMES] [varchar](30) NULL,[VOLUMES] [varchar](30) NULL,[CRO] [smallint] NULL,[USARATEIOVALORFIN] [FLOAT] NULL,[RECIBONFESERIE] [varchar](5) NULL,[CODCOLCFOORIGEM] [FLOAT] NULL,[CODCFOORIGEM] [varchar](25) NULL,[VALORCTRCARATEAR] [FLOAT] NULL,[CODCOLCFOAUX] [smallint] NULL,[VRBASEINSSOUTRAEMPRESA] [FLOAT] NULL,[IDCEICFO] [int] NULL,[CHAVEACESSONFE] [varchar](44) NULL,[VLRSECCAT] [FLOAT] NULL,[VLRDESPACHO] [FLOAT] NULL,[VLRPEDAGIO] [FLOAT] NULL,[VLRFRETEOUTROS] [FLOAT] NULL,[ABATIMENTONAOTRIB] [FLOAT] NULL, " & _
        "   [RATEIOCCUSTODEPTO] [FLOAT] NULL,[VALORRATEIOLAN] [FLOAT] NULL,[CODCOLCFOTRANSFAT] [FLOAT] NULL,[CODCFOTRANSFAT] [varchar](25) NULL,[CODUSUARIOAPROVADESC] [varchar](20) NULL,[IDINTEGRACAO] [varchar](100) NULL) ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF  " & _
        "ALTER TABLE [dbo].[TMOV] ADD [STATUSANTERIOR] [varchar](1) NULL ALTER TABLE [dbo].[TMOV] ADD [VALORBRUTOORIG] [FLOAT] NULL ALTER TABLE [dbo].[TMOV] ADD [VALORLIQUIDOORIG] [FLOAT] NULL ALTER TABLE [dbo].[TMOV] ADD [VALOROUTROSORIG] [FLOAT] NULL ALTER TABLE [dbo].[TMOV] ADD [VALORRATEIOLANORIG] [FLOAT] NULL ALTER TABLE [dbo].[TMOV] ADD [IDOPERACAO] [int] NULL ALTER TABLE [dbo].[TMOV] ADD [DATAPROCESSAMENTO] [DATETIME] NULL ALTER TABLE [dbo].[TMOV] ADD [IDNATFRETE] [int] NULL ALTER TABLE [dbo].[TMOV] ADD [RECCREATEDBY] [varchar](50) NULL " & _
        "ALTER TABLE [dbo].[TMOV] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TMOV] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TMOV] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[TMOV] ADD  CONSTRAINT [PKTMOV] PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[IDMOV] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    
        'TITMMOV
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TITMMOV]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[IDMOV] [int] NOT NULL,[NSEQITMMOV] [int] NOT NULL,[NUMEROSEQUENCIAL] [smallint] NOT NULL,[IDPRD] [int] NULL,[CODTIP] [varchar](10) NULL,[QUANTIDADE] [FLOAT] NULL,[PRECOUNITARIO] [decimal](21, 10) NULL,[PRECOTABELA] [FLOAT] NULL,[PERCENTUALDESC] [FLOAT] NULL,[VALORDESC] [FLOAT] NULL,[PERCENTUALDESP] [FLOAT] NULL,[VALORDESP] [FLOAT] NULL,[DATAEMISSAO] [datetime] NULL,[CODMEN] [varchar](5) NULL,[NUMEROTRIBUTOS] [smallint] NULL, " & _
        "   [CODTB1FAT] [varchar](10) NULL,[CODTB2FAT] [varchar](10) NULL,[CODTB3FAT] [varchar](10) NULL,[CODTB4FAT] [varchar](10) NULL,[CODTB5FAT] [varchar](10) NULL,[CODTB1FLX] [varchar](25) NULL,[CODTB2FLX] [varchar](25) NULL,[CODTB3FLX] [varchar](25) NULL,[CODTB4FLX] [varchar](25) NULL,[CODTB5FLX] [varchar](25) NULL,[CAMPOLIVRE] [varchar](15) NULL,[CODUND] [varchar](5) NULL,[QUANTIDADEARECEBER] [FLOAT] NULL,[CODNAT] [varchar](10) NULL,[CODCPG] [varchar](5) NULL, " & _
        "   [DATAENTREGA] [datetime] NULL,[PRATELEIRA] [varchar](15) NULL,[IDCNT] [int] NULL,[NSEQITMCNT] [smallint] NULL,[DATAINIFAT] [datetime] NULL,[DATAFIMFAT] [datetime] NULL,[FLAGEFEITOSALDO] [smallint] NULL,[VALORUNITARIO] [FLOAT] NULL,[VALORFINANCEIRO] [FLOAT] NULL,[IMPRIMEMOV] [smallint] NULL,[CODCCUSTO] [varchar](25) NULL,[FLAGREPASSE] [smallint] NULL,[ALIQORDENACAO] [FLOAT] NULL,[QUANTIDADEORIGINAL] [FLOAT] NULL,[IDNAT] [int] NULL,[FLAG] [smallint] NULL, " & _
        "   [CHAPA] [VARCHAR] (20) NULL,[INICIO] [datetime] NULL,[TERMINO] [datetime] NULL,[PREVINICIO] [datetime] NULL,[STATUS] [char](1) NULL,[BLOCK] [smallint] NULL,[FLAGREFATURAMENTO] [smallint] NULL,[IDCNTDESTINO] [int] NULL,[NSEQITMCNTDEST] [smallint] NULL,[FATORCONVUND] [FLOAT] NULL,[IDPRJ] [int] NULL,[IDTRF] [int] NULL,[VALORTOTALITEM] [decimal](21, 10) NULL,[VALORCODIGOPRD] [varchar](60) NULL,[TIPOCODIGOPRD] [smallint] NULL,[QTDUNDPEDIDO] [FLOAT] NULL, " & _
        "   [TRIBUTACAOECF] [varchar](10) NULL,[CODFILIAL] [smallint] NULL,[CODDEPARTAMENTO] [varchar](25) NULL,[IDPRDCOMPOSTO] [int] NULL,[QUANTIDADESEPARADA] [FLOAT] NULL,[PERCENTCOMISSAO] [FLOAT] NULL,[INDICENCM] [char](1) NULL,[NCM] [varchar](14) NULL,[CODRPR] [varchar](15) NULL,[COMISSAOREPRES] [FLOAT] NULL,[NSEQITMCNTMEDICAO] [smallint] NULL,[VALORESCRITURACAO] [FLOAT] NULL,[VALORFINPEDIDO] [FLOAT] NULL,[VALORFRETECTRC] [FLOAT] NULL,[VALOROPFRM1] [FLOAT] NULL, " & _
        "   [VALOROPFRM2] [FLOAT] NULL,[IDOBJOFICINA] [varchar](20) NULL,[PRECOEDITADO] [FLOAT] NULL,[QTDEVOLUMEUNITARIO] [smallint] NULL,[IDGRD] [int] NULL,[CODVEN1] [varchar](16) NULL,[CODLOCALBN] [varchar](40) NULL,[REGISTROEXPORTACAO] [varchar](12) NULL,[DATARE] [datetime] NULL,[PRECOTOTALEDITADO] [INT] NULL,[CST] [varchar](3) NULL,[VALORDESCCONDICONALITM] [FLOAT] NULL,[VALORDESPCONDICIONALITM] [FLOAT] NULL,[DATAORCAMENTO] [datetime] NULL,[CODTBORCAMENTO] [varchar](40) NULL, " & _
        "   [RATEIOFRETE] [FLOAT] NULL,[RATEIOSEGURO] [FLOAT] NULL,[RATEIODESC] [FLOAT] NULL,[RATEIODESP] [FLOAT] NULL,[RATEIOEXTRA1] [FLOAT] NULL,[RATEIOEXTRA2] [FLOAT] NULL,[RATEIOFRETECTRC] [FLOAT] NULL,[RATEIODEDMAT] [FLOAT] NULL,[RATEIODEDSUB] [FLOAT] NULL,[RATEIODEDOUT] [FLOAT] NULL,[IDCLASSIFENERGIACOMUNIC] [int] NULL,[VALORUNTORCAMENTO] [FLOAT] NULL,[VALSERVICONFE] [FLOAT] NULL,[CODLOC] [varchar](15) NULL,[VALORBEM] [FLOAT] NULL,[VALORLIQUIDO] [FLOAT] NULL, " & _
        "   [CODIGOCODIF] [varchar](21) NULL,[CODMUNSERVICO] [varchar](20) NULL,[CODETDMUNSERV] [varchar](2) NULL,[RATEIOCCUSTODEPTO] [FLOAT] NULL,[CUSTOREPOSICAO] [FLOAT] NULL,[CUSTOREPOSICAOB] [FLOAT] NULL,[VALORFINTERCEIROS] [FLOAT] NULL,[VALORFINANCGERENCIAL] [FLOAT] NULL,[CODIGOSERVICO] [varchar](15) NULL,[VALORUNITGERENCIAL] [FLOAT] NULL,[IDINTEGRACAO] [varchar](100) NULL,[IDTABPRECO] [int] NULL,[VALORBRUTOITEM] [decimal](21, 10) NULL,[VALORBRUTOITEMORIG] [decimal](21, 10) NULL, " & _
        "   [CODCOLTBORCAMENTO] [FLOAT] NULL,[CODPUBLIC] [int] NULL,[QUANTIDADETOTAL] [FLOAT] NULL,[PRODUTOSUBSTITUTO] [FLOAT] NOT NULL)  " & _
        "ON [PRIMARY] SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[TITMMOV] ADD [CODTBGRUPOORC] [varchar](40) NULL ALTER TABLE [dbo].[TITMMOV] ADD [PRECOUNITARIOSELEC] [int] NULL ALTER TABLE [dbo].[TITMMOV] ADD [VALORRATEIOLAN] [FLOAT] NULL ALTER TABLE [dbo].[TITMMOV] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TITMMOV] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TITMMOV] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TITMMOV] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[TITMMOV] ADD [IDCNTOP] [int] NULL " & _
        "ALTER TABLE [dbo].[TITMMOV] ADD [QUANTIDADECONCLUIDA] [FLOAT] NULL ALTER TABLE [dbo].[TITMMOV] ADD [DATAFATCONTRATO] [DATETIME] NULL ALTER TABLE [dbo].[TITMMOV] ADD  CONSTRAINT [PKTITMMOV]  " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[IDMOV] ASC,[NSEQITMMOV] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'TPRDLOC
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TPRDLOC]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[CODFILIAL] [smallint] NOT NULL,[CODLOC] [varchar](15) NOT NULL,[IDPRD] [int] NOT NULL,[SALDOFISICO1] [DECIMAL] (15,4) NULL,[SALDOFISICO2] [DECIMAL] (15,4) NULL,[SALDOFISICO3] [DECIMAL] (15,4) NULL,[SALDOFISICO4] [DECIMAL] (15,4) NULL,[SALDOFISICO5] [DECIMAL] (15,4) NULL,[SALDOFISICO6] [DECIMAL] (15,4) NULL,[SALDOFISICO7] [DECIMAL] (15,4) NULL,[SALDOFISICO8] [DECIMAL] (15,4) NULL,[SALDOFISICO9] [DECIMAL] (15,4) NULL,[SALDOFISICO10] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO1] [DECIMAL] (15,4) NULL, " & _
        "   [SALDOFINANCEIRO2] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO3] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO4] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO5] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO6] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO7] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO8] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO9] [DECIMAL] (15,4) NULL,[SALDOFINANCEIRO10] [DECIMAL] (15,4) NULL,[CUSTOMEDIO] [DECIMAL] (15,4) NULL,[CUSTOREPOSICAO] [DECIMAL] (15,4) NULL,[CUSTOREPOSICAOB] [DECIMAL] (15,4) NULL,[CUSTOUNITARIO] [DECIMAL] (15,4) NULL, " & _
        "   [DATACUSTOMEDIO] [datetime] NULL,[DATACUSTOREPOSICAO] [datetime] NULL,[DATACUSTOREPOSICAOB] [datetime] NULL,[DATACUSTOUNITARIO] [datetime] NULL,[PRATELEIRA] [varchar](15) NULL,[VENDIDODETERCEIROS] [DECIMAL] (15,4) NULL " & _
        ") ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[TPRDLOC] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TPRDLOC] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TPRDLOC] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TPRDLOC] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[TPRDLOC] ADD  CONSTRAINT [PKTPRDLOC]  " & _
        "PRIMARY KEY CLUSTERED ([IDPRD] ASC,[CODLOC] ASC,[CODFILIAL] ASC,[CODCOLIGADA] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'TMOVRELAC
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TMOVRELAC]( " & _
        "   [IDMOVORIGEM] [INT] NOT NULL,[CODCOLORIGEM] [SMALLINT] NOT NULL,[IDMOVDESTINO] [INT] NOT NULL,[CODCOLDESTINO] [SMALLINT] NOT NULL,[TIPORELAC] [varchar](1) NOT NULL,[IDPROCESSO] [int] NULL " & _
        ") ON [PRIMARY] SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[TMOVRELAC] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TMOVRELAC] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TMOVRELAC] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TMOVRELAC] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[TMOVRELAC] ADD  CONSTRAINT [PKTMOVRELAC]  " & _
        "PRIMARY KEY CLUSTERED ([IDMOVORIGEM] ASC,[CODCOLORIGEM] ASC,[IDMOVDESTINO] ASC,[CODCOLDESTINO] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'TVEN
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TVEN]( " & _
        "   [CODCOLIGADA] [SMALLINT] NOT NULL,[CODVEN] [varchar](16) NOT NULL,[NOME] [varchar](80) NULL,[CARGO] [varchar](30) NULL,[CODFILIAL] [smallint] NULL,[CODLOC] [varchar](15) NULL,[COMISSAO1] [DECIMAL] (15,4) NULL,[COMISSAO2] [DECIMAL] (15,4) NULL,[COMISSAO3] [DECIMAL] (15,4) NULL,[CODPESSOA] [int] NULL,[VENDECOMPRA] [smallint] NULL,[CODUSUARIO] [varchar](20) NULL,[SENHA] [varchar](80) NULL,[INATIVO] [INT] NULL,[PFVENDEDOR] [INT] NULL, " & _
        "   [PFCAIXA] [INT] NULL,[PFSUPERVISOR] [INT] NULL,[PFGERENTE] [INT] NULL,[IDFUNCIONARIO] [int] NOT NULL,[COMISSAO4] [DECIMAL] (15,4) NULL,[DESCMAXIMO] [DECIMAL] (15,4) NULL " & _
        ") ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[TVEN] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TVEN] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TVEN] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TVEN] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[TVEN] ADD  CONSTRAINT [PKTVEN]  " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CODVEN] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'TVENCOMPL
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TVENCOMPL]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[CODVEN] [varchar](16) NOT NULL " & _
        ") ON [PRIMARY]SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[TVENCOMPL] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TVENCOMPL] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TVENCOMPL] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TVENCOMPL] ADD [RECMODIFIEDON] [datetime] NULL " & _
        "SET ANSI_PADDING ON ALTER TABLE [dbo].[TVENCOMPL] ADD [NUMCALCADO] [varchar](2) NULL ALTER TABLE [dbo].[TVENCOMPL] ADD [SITEMPRESTIMO] [varchar](10) NULL ALTER TABLE [dbo].[TVENCOMPL] ADD [MOTBLOQUEIO] [varchar](60) NULL ALTER TABLE [dbo].[TVENCOMPL] ADD  CONSTRAINT [PKTVENCOMPL] " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CODVEN] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'PFUNCAO
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[PFUNCAO]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[CODIGO] [varchar](10) NOT NULL,[NOME] [varchar](40) NULL,[NUMPONTOS] [DECIMAL] (15,4) NULL,[CBO] [varchar](8) NULL,[CARGO] [varchar](16) NULL,[INATIVA] [smallint] NULL,[ATIVTRANSP] [smallint] NULL,[DESCRICAO] [TEXT] NULL,[FAIXASALARIAL] [varchar](16) NULL,[LIMITEFUNC] [int] NULL,[VERBAQUADROVAGAS] [DECIMAL] (15,4) NULL,[PERCQUADROVAGAS] [DECIMAL] (15,4) NULL,[DATAULTIMAREVISAO] [datetime] NULL,[NUMREVISAO] [varchar](30) NULL,[CBO2002] [varchar](10) NULL, " & _
        "   [CODTABELA] [varchar](10) NULL,[CODPERFILCAND] [varchar](15) NULL,[ID] [int] IDENTITY(1,1) NOT NULL,[BENEFPONTOS] [int] NULL,[OBJETIVO] [TEXT] NULL,[DESCRICAOPPP] [TEXT] NULL " & _
        ") ON [PRIMARY]SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[PFUNCAO] ADD [EXIBEORGANOGRAMA] [CHAR] (1) NULL ALTER TABLE [dbo].[PFUNCAO] ADD [CODFUNCAOCHEFIA] [varchar](10) NULL ALTER TABLE [dbo].[PFUNCAO] ADD [JORNADAREF] [DECIMAL] (15,4) NULL ALTER TABLE [dbo].[PFUNCAO] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[PFUNCAO] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[PFUNCAO] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[PFUNCAO] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[PFUNCAO] ADD  CONSTRAINT [PKPFUNCAO]  " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CODIGO] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'PSECAO
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[PSECAO]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[CODIGO] [varchar](35) NOT NULL,[DESCRICAO] [varchar](60) NULL,[CGC] [varchar](20) NULL,[CGCANTERIOR] [varchar](20) NULL,[CEI] [varchar](20) NULL,[FPAS] [varchar](3) NULL,[SAT] [varchar](12) NULL,[ATIVECONOMICA] [varchar](7) NULL,[INSCRESTADUAL] [varchar](15) NULL,[INSCRMUNICIPAL] [varchar](15) NULL,[RUA] [varchar](100) NULL,[NUMERO] [varchar](8) NULL,[COMPLEMENTO] [TEXT] NULL,[BAIRRO] [VARCHAR] (50) NULL,[ESTADO] [varchar](2) NULL, " & _
        "   [CIDADE] [VARCHAR] (50) NULL,[CEP] [VARCHAR] (10) NULL,[PAIS] [varchar](16) NULL,[TELEFONE] [varchar](15) NULL,[CODIGOCEF] [varchar](14) NULL,[NAOEMPREGPROPR] [smallint] NULL,[PREFIXORAIS] [varchar](2) NULL,[CATEGORIA] [varchar](2) NULL,[INTEGRCONTABIL] [varchar](22) NULL,[INTEGRGERENCIAL] [varchar](22) NULL,[NROFILIALCONT] [smallint] NULL,[NROCENCUSTOCONT] [varchar](25) NULL,[CODTERCEIROSINSS] [varchar](4) NULL,[PERCENTTERCEIROS] [DECIMAL] (15,4) NULL, " & _
        "   [PERCENTACIDTRAB] [DECIMAL] (15,4) NULL,[FATURAMENTOBRUTO] [DECIMAL] (15,4) NULL,[VALORFRETE] [DECIMAL] (15,4) NULL,[COMPLEMENTOGRPS1] [varchar](58) NULL,[COMPLEMENTOGRPS2] [varchar](58) NULL,[COMPLEMENTOGRPS3] [varchar](58) NULL,[CODUNIDENTREGA] [varchar](6) NULL,[CONTATO] [varchar](20) NULL,[RAMAL] [varchar](4) NULL,[VTCODDEPTO] [varchar](6) NULL,[INIPERRECPROPR1] [smallint] NULL,[FIMPERRECPROPR1] [smallint] NULL,[PROPRANTES5DIA1] [SMALLINT] NULL, " & _
        "   [INIPERRECPROPR2] [smallint] NULL,[FIMPERRECPROPR2] [smallint] NULL,[PROPRANTES5DIA2] [SMALLINT] NULL,[INIPERRECCENTR1] [smallint] NULL,[FIMPERRECCENTR1] [smallint] NULL,[CENTRANTES5DIA1] [SMALLINT] NULL,[INIPERRECCENTR2] [smallint] NULL,[FIMPERRECCENTR2] [smallint] NULL,[CENTRANTES5DIA2] [SMALLINT] NULL,[SECAOCENTRALIZ] [varchar](35) NULL,[CONTRIBSESIESENAI] [SMALLINT] NULL,[DISTRIBPETROLEO] [SMALLINT] NULL,[PESSOAFISICA] [SMALLINT] NULL, " & _
        "   [SECAODESATIVADA] [SMALLINT] NULL,[IDENTIFICACAOCGC] [SMALLINT] NULL,[ENDERECOALTEROU] [SMALLINT] NULL,[ENDERECOPAGTO] [varchar](48) NULL,[VTIDENTPEDIDO] [smallint] NULL,[VTIDENTPERSONAL] [smallint] NULL,[VTCODCLIENTE] [varchar](6) NULL,[VTCODLOCAL] [varchar](10) NULL,[VTCODSECFATURAM] [varchar](35) NULL,[VTCODSECPEDIDO] [varchar](35) NULL,[VTCODSECCOBRANCA] [varchar](35) NULL,[VTCODSECENTREGA] [varchar](35) NULL,[VTCODCENTROCUSTO] [varchar](25) NULL, " & _
        "   [CAUSAMUDANCACGC] [smallint] NULL,[CODMUNICIPIO] [varchar](7) NULL,[MESDATABASE] [smallint] NULL,[NATUREZAJURIDICA] [varchar](4) NULL,[CODCALENDARIO] [varchar](16) NULL,[PREFIXOINSCRFGTS] [varchar](2) NULL,[PRIMEIRADECLCAGED] [smallint] NULL,[ENCERRAMENTO] [smallint] NULL,[CODFILIAL] [smallint] NOT NULL,[CODDEPTO] [varchar](25) NOT NULL,[LIMITEFUNC] [int] NULL,[OPTASIMPLES] [smallint] NULL,[MUDOUCNAE] [SMALLINT] NULL,[CHAPACHEFE] [VARCHAR] (20) NULL, " & _
        "   [PERCENT15ANOSGRPS] [smallint] NULL,[PERCENT20ANOSGRPS] [smallint] NULL,[PERCENT25ANOSGRPS] [smallint] NULL,[ALTERACAOCAGED] [smallint] NULL,[CODPAGTOGPS] [varchar](5) NULL,[PERCISENCAOFILANTROPIA] [DECIMAL] (15,4) NULL,[PARTICIPAPAT] [smallint] NULL,[PORTEEMPRESA] [smallint] NULL,[CODDEPTOCONT] [varchar](25) NULL,[DDD] [varchar](4) NULL,[VERBAQUADROVAGAS] [DECIMAL] (15,4) NULL,[PERCQUADROVAGAS] [DECIMAL] (15,4) NULL,[ISENTOCONTRIBSOCIAL] [DECIMAL] (15,4) NULL, " & _
        "   [VINCPAT5SAL] [int] NULL,[VINCPATMAIOR5SAL] [int] NULL,[PORCSERVPROP] [smallint] NULL,[PORCADMCOZINHA] [smallint] NULL,[PORCREFEICAOCONV] [smallint] NULL,[PORCREFEICAOTRANSP] [smallint] NULL,[PORCCESTAALIMENTO] [smallint] NULL,[PORCALIMCONVENIO] [smallint] NULL,[DTENCERRATIV] [datetime] NULL,[CODPLANO] [varchar](16) NULL,[DESCRICAOPPP] [TEXT] NULL,[CODTABELA] [varchar](10) NULL,[EMAIL] [varchar](45) NULL,[CNAERAIS] [varchar](5) NULL,[VALORINSSACUMULADO] [DECIMAL] (15,4) NULL, " & _
        "   [VALORENTIDADESACUMULADO] [DECIMAL] (15,4) NULL,[IDMEMOAMBTRAB] [int] NULL,[VISIVELORGANOGRAMA] [CHAR] (1) NULL,[LOCALIDADE] [varchar](40) NULL,[CAPITALSOCEMP] [DECIMAL] (15,4) NULL,[CAPITALSOCESTAB] [DECIMAL] (15,4) NULL,[CODIGOPAI] [varchar](35) NULL,[ID] [int] IDENTITY(1,1) NOT NULL,[BENEFPONTOS] [int] NULL) ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF ALTER TABLE [dbo].[PSECAO] ADD [CODPAGTOGPSTERCEIROS] [varchar](5) NULL ALTER TABLE [dbo].[PSECAO] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[PSECAO] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[PSECAO] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[PSECAO] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[PSECAO] ADD [TPLOTACAO] [varchar](2) NULL ALTER TABLE [dbo].[PSECAO] ADD [CNO] [varchar](20) NULL " & _
        "ALTER TABLE [dbo].[PSECAO] ADD [CODTIPORUA] [smallint] NULL ALTER TABLE [dbo].[PSECAO] ADD [TPINSCCONTRATANTE] [smallint] NULL ALTER TABLE [dbo].[PSECAO] ADD [NROINSCCONTRATANTE] [varchar](20) NULL ALTER TABLE [dbo].[PSECAO] ADD [TPINSCPROPRIETARIO] [smallint] NULL ALTER TABLE [dbo].[PSECAO] ADD [NROINSCPROPRIETARIO] [varchar](20) NULL ALTER TABLE [dbo].[PSECAO] ADD [CAEPF] [varchar](20) NULL ALTER TABLE [dbo].[PSECAO] ADD  CONSTRAINT [PKPSECAO]  " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CODIGO] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'TPRODUTO
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TPRODUTO]( " & _
        "   [CODCOLPRD] [INT] NOT NULL,[IDPRD] [int] NOT NULL,[CODIGOPRD] [varchar](30) NULL,[NOMEFANTASIA] [varchar](100) NULL,[CODIGOREDUZIDO] [varchar](10) NULL,[ULTIMONIVEL] [smallint] NULL,[TIPO] [varchar](1) NULL,[DESCRICAO] [varchar](240) NULL,[DESCRICAOAUX] [varchar](240) NULL,[CODIGOAUXILIAR] [varchar](20) NULL,[REFERENCIACCF] [varchar](1) NULL,[NUMEROCCF] [varchar](14) NULL,[REFERENCIACP] [smallint] NULL,[DESCRICAOCP] [varchar](40) NULL,[PESOLIQUIDO] [DECIMAL] (15,4) NULL,[PESOBRUTO] [DECIMAL] (15,4) NULL, " & _
        "   [COMPRIMENTO] [DECIMAL] (15,4) NULL,[ESPESSURA] [DECIMAL] (15,4) NULL,[LARGURA] [DECIMAL] (15,4) NULL,[COR] [varchar](15) NULL,[OBSERVACAO] [varchar](10) NULL,[DTCADASTRAMENTO] [datetime] NULL,[CAMPOLIVRE] [int] NULL,[CAMPOLIVRE2] [varchar](5) NULL,[CAMPOLIVRE3] [varchar](15) NULL,[IDPRODUTORELAC] [int] NULL,[TEMPO] [DECIMAL] (15,4) NULL,[INATIVO] [CHAR] (1) NULL,[PESAVEL] [smallint] NULL,[DATAULTALTERACAO] [datetime] NULL,[IDIMAGEM] [int] NULL,[DIAMETRO] [DECIMAL] (15,4) NULL,[CODUSUARIO] [varchar](20) NULL, " & _
        "   [TEMPOVALIDADE] [int] NULL,[USUARIOCRIACAO] [varchar](20) NULL,[QTDEVOLUME] [smallint] NULL,[CODIGOEX] [varchar](3) NULL,[SERVICOPRODUTORMOFFICINA] [smallint] NULL,[PRODUTOEPI] [smallint] NULL,[PRODUTOBASE] [CHAR] (1) NULL,[NUMEROTRIBUTOS] [smallint] NULL,[EPERIODICO] [smallint] NULL,[BLOCK] [smallint] NULL,[DATAEXTRA1] [datetime] NULL,[DATAEXTRA2] [datetime] NULL,[CONTROLADOPORLOTE] [int] NULL,[MASCARANUMSERIE] [varchar](30) NULL,[USANUMSERIE] [int] NULL,[VALIDADEMINIMA] [int] NULL,[ID] [int] IDENTITY(1,1) NOT NULL, " & _
        "   [RECCREATEDBY] [varchar](50) NULL,[RECCREATEDON] [datetime] NULL,[RECMODIFIEDBY] [varchar](50) NULL,[RECMODIFIEDON] [datetime] NULL,[IDNCM] [int] NULL, " & _
        "CONSTRAINT [PKTPRD] PRIMARY KEY CLUSTERED ([IDPRD] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
        
        'TPRODUTODEF
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TPRODUTODEF]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[IDPRD] [int] NOT NULL,[CODFAB] [varchar](15) NULL,[NUMNOFABRIC] [varchar](100) NULL,[PRECO1] [DECIMAL] (15,4) NULL,[PRECO2] [DECIMAL] (15,4) NULL,[PRECO3] [DECIMAL] (15,4) NULL,[PRECO4] [DECIMAL] (15,4) NULL,[PRECO5] [DECIMAL] (15,4) NULL,[DATABASEPRECO1] [datetime] NULL,[DATABASEPRECO2] [datetime] NULL,[DATABASEPRECO3] [datetime] NULL,[DATABASEPRECO4] [datetime] NULL,[DATABASEPRECO5] [datetime] NULL,[CODMOEPRECO1] [varchar](10) NULL, " & _
        "   [CODMOEPRECO2] [varchar](10) NULL,[CODMOEPRECO3] [varchar](10) NULL,[CODMOEPRECO4] [varchar](10) NULL,[CODMOEPRECO5] [varchar](10) NULL,[CODUNDCONTROLE] [varchar](5) NULL,[CODCPG] [varchar](5) NULL,[MARGEMLUCROFISC] [DECIMAL] (15,4) NULL,[FATORREDUCAOICMS] [DECIMAL] (15,4) NULL,[ESTOQUEMINIMO1] [DECIMAL] (15,4) NULL,[ESTOQUEMAXIMO1] [DECIMAL] (15,4) NULL,[ESTOQUEMINIMO2] [DECIMAL] (15,4) NULL,[ESTOQUEMAXIMO2] [DECIMAL] (15,4) NULL,[ESTOQUEMINIMO3] [DECIMAL] (15,4) NULL, " & _
        "   [ESTOQUEMAXIMO3] [DECIMAL] (15,4) NULL,[PONTODEPEDIDO1] [DECIMAL] (15,4) NULL,[PONTODEPEDIDO2] [DECIMAL] (15,4) NULL,[PONTODEPEDIDO3] [DECIMAL] (15,4) NULL,[PERCENTCOMISSAO] [DECIMAL] (15,4) NULL,[DESCONTOCOMPRA] [DECIMAL] (15,4) NULL,[DESCONTOVENDA] [DECIMAL] (15,4) NULL,[CUSTOUNITARIO] [DECIMAL] (15,4) NULL,[DTCUSTOUNITARIO] [datetime] NULL,[CODTIP] [varchar](10) NULL,[CODTB1FAT] [varchar](10) NULL,[CODTB2FAT] [varchar](10) NULL,[CODTB3FAT] [varchar](10) NULL, " & _
        "   [CODTB4FAT] [varchar](10) NULL,[CODTB5FAT] [varchar](10) NULL,[MARGEMBRUTALUCRO] [DECIMAL] (15,4) NULL,[SALDOGERALFISICO] [DECIMAL] (15,4) NULL,[PERCENTCOMISSAO2] [DECIMAL] (15,4) NULL,[CODUNDCOMPRA] [varchar](5) NULL,[CODUNDVENDA] [varchar](5) NULL,[RECALCCUSTOMEDIO] [smallint] NULL,[CUSTOMEDIO] [DECIMAL] (15,4) NULL,[DATACUSTOMEDIO] [datetime] NULL,[DTULTIMACOMPRA] [datetime] NULL,[TIPOCONTA] [varchar](1) NULL,[SALDOGERALFINANC] [DECIMAL] (15,4) NULL, " & _
        "   [DTULTIMACOMPRAB] [datetime] NULL,[CUSTOREPOSICAOB] [DECIMAL] (15,4) NULL,[PERCENTCOMISSAO3] [DECIMAL] (15,4) NULL,[CUSTOREPOSICAO] [DECIMAL] (15,4) NULL,[LOCALDESCARGA] [varchar](30) NULL,[CODCOLCONTAGER] [SMALLINT] NULL,[CODCONTAGER] [SMALLINT] NULL,[USANUMDECPRECO] [SMALLINT] NULL,[NUMDECPRECO] [smallint] NULL,[CLASSEFISCALECF] [smallint] NULL,[MULTIPLOPRD] [DECIMAL] (15,4) NULL,[CODTIPOAPL] [SMALLINT] NULL,[CODDIEF] [varchar](7) NULL,[TOLERANCIASUP] [DECIMAL] (15,4) NULL, " & _
        "   [TOLERANCIAINF] [DECIMAL] (15,4) NULL,[CODGRUPO] [varchar](10) NULL,[TRIBUTACAOECF] [varchar](10) NULL,[TIPOCALCULOCUSTO] [smallint] NULL,[CUSTOPADRAO] [int] NULL,[REPASSEFABRIC] [DECIMAL] (15,4) NULL,[INVENTARIOFISCAL] [smallint] NULL,[CODGRUPOBEM] [SMALLINT] NULL,[GRPFATURAMENTO] [varchar](10) NULL,[TOLINFPRECO] [DECIMAL] (15,4) NULL,[TOLSUPPRECO] [DECIMAL] (15,4) NULL,[MULTIPLOQTDECOMPRADA] [DECIMAL] (15,4) NULL,[MULTIPLOPRDVENDA] [DECIMAL] (15,4) NULL,[CODTBORCAMENTO] [varchar](40) NULL, " & _
        "   [IDPRDFISCALS] [int] NULL,[IDPRDFISCALE] [int] NULL,[CODCOLTBORCAMENTO] [SMALLINT] NULL,[RECALCSALDO1] [smallint] NULL,[RECALCSALDO2] [smallint] NULL,[RECALCSALDO3] [smallint] NULL,[RECALCSALDO4] [smallint] NULL,[RECALCSALDO5] [smallint] NULL,[RECALCSALDO6] [smallint] NULL,[RECALCSALDO7] [smallint] NULL,[RECALCSALDO8] [smallint] NULL,[RECALCSALDO9] [smallint] NULL,[RECALCSALDO10] [smallint] NULL,[DATAPRIMEIRAALT] [datetime] NULL,[IDGRD] [int] NULL,[CODGRD] [varchar](20) NULL, " & _
        "   [CODCOLUNA] [varchar](3) NULL,[CODLINHA] [varchar](3) NULL,[RECCREATEDBY] [varchar](50) NULL,[RECCREATEDON] [datetime] NULL,[RECMODIFIEDBY] [varchar](50) NULL,[RECMODIFIEDON] [datetime] NULL, " & _
        "CONSTRAINT [PKTPRDDEF] PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[IDPRD] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
        
        'OFVENCPLANOMANUT
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[OFVENCPLANOMANUT]( " & _
        "   [CODCOLIGADA] [INT] NOT NULL,[IDOBJOF] [varchar](20) NOT NULL,[IDPLANO] [int] NOT NULL,[DATAATUALIZACAO] [datetime] NULL,[DATAVENCIMENTO] [datetime] NULL,[CODUSUARIO] [varchar](20) NULL) ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[OFVENCPLANOMANUT] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[OFVENCPLANOMANUT] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[OFVENCPLANOMANUT] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[OFVENCPLANOMANUT] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[OFVENCPLANOMANUT] ADD  CONSTRAINT [PKOFVENCPLANOMANUT]  " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[IDOBJOF] ASC,[IDPLANO] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'OFPLANOMANUT
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[OFPLANOMANUT]( " & _
        "   [IDPLANO] [int] NOT NULL,[DESCRICAO] [varchar](60) NOT NULL,[FREQUENCIA] [int] NOT NULL,[TIPOFREQ] [smallint] NULL,[IDTIPOOBJ] [int] NULL,[CODMODELO] [smallint] NULL,[CODSUBMODELO] [smallint] NULL,[OBS] [TEXT] NULL,[CAMPOLIVREC1] [varchar](20) NULL,[CAMPOLIVREC2] [varchar](20) NULL,[CAMPOLIVREC3] [varchar](20) NULL,[CAMPOLIVREC4] [varchar](20) NULL,[DIA] [DECIMAL] (15,4) NULL,[INDICADORUSO1] [DECIMAL] (15,4) NULL,[INDICADORUSO2] [DECIMAL] (15,4) NULL,[INDICADORUSO3] [DECIMAL] (15,4) NULL, " & _
        "   [INDICADORUSO4] [DECIMAL] (15,4) NULL,[INDICADORUSO5] [DECIMAL] (15,4) NULL,[ATIVO] [CHAR] (1) NOT NULL) ON [PRIMARY]  " & _
        "SET ANSI_PADDING OFF ALTER TABLE [dbo].[OFPLANOMANUT] ADD [TIPOEXECUCAO] [char](1) NOT NULL ALTER TABLE [dbo].[OFPLANOMANUT] ADD [ANOMODELOINI] [smallint] NULL ALTER TABLE [dbo].[OFPLANOMANUT] ADD [ANOMODELOFIM] [smallint] NULL ALTER TABLE [dbo].[OFPLANOMANUT] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[OFPLANOMANUT] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[OFPLANOMANUT] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[OFPLANOMANUT] ADD [RECMODIFIEDON] [datetime] NULL " & _
        "ALTER TABLE [dbo].[OFPLANOMANUT] ADD  CONSTRAINT [PKOFPLANOMANUT] PRIMARY KEY CLUSTERED ([IDPLANO] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
        
        'TLOC
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TLOC]( " & _
        "   [CODFILIAL] [smallint] NOT NULL,[CODLOC] [varchar](15) NOT NULL,[NOME] [varchar](40) NULL,[NIVELESTOQUE] [smallint] NULL,[CODCOLIGADA] [SMALLINT] NOT NULL,[IDUNDNEGOCIO] [int] NULL) ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF ALTER TABLE [dbo].[TLOC] ADD [RUA] [varchar](100) NULL " & _
        "SET ANSI_PADDING ON ALTER TABLE [dbo].[TLOC] ADD [COMPLEMENTO] [VARCHAR] (60) NULL " & _
        "ALTER TABLE [dbo].[TLOC] ADD [BAIRRO] [VARCHAR] (80) NULL SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[TLOC] ADD [CIDADE] [VARCHAR] (32) NULL ALTER TABLE [dbo].[TLOC] ADD [CEP] [VARCHAR] (9) NULL ALTER TABLE [dbo].[TLOC] ADD [CONTATO] [varchar](40) NULL ALTER TABLE [dbo].[TLOC] ADD [DDD] [varchar](4) NULL ALTER TABLE [dbo].[TLOC] ADD [EMAIL] [varchar](60) NULL ALTER TABLE [dbo].[TLOC] ADD [CODETD] [varchar](2) NULL " & _
        "ALTER TABLE [dbo].[TLOC] ADD [FAX] [varchar](15) NULL ALTER TABLE [dbo].[TLOC] ADD [NUMERO] [varchar](8) NULL ALTER TABLE [dbo].[TLOC] ADD [PAIS] [varchar](16) NULL ALTER TABLE [dbo].[TLOC] ADD [TELEFONE] [varchar](15) NULL ALTER TABLE [dbo].[TLOC] ADD [INATIVO] [smallint] NULL ALTER TABLE [dbo].[TLOC] ADD [RECCREATEDBY] [varchar](50) NULL " & _
        "ALTER TABLE [dbo].[TLOC] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[TLOC] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[TLOC] ADD [RECMODIFIEDON] [datetime] NULL " & _
        "ALTER TABLE [dbo].[TLOC] ADD  CONSTRAINT [PKTLOC] PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CODFILIAL] ASC,[CODLOC] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'GCOLIGADA
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[GCOLIGADA]( " & _
        "   [CODCOLIGADA] [SMALLINT] NOT NULL,[NOMEFANTASIA] [varchar](60) NULL,[CGC] [varchar](20) NULL,[NOME] [varchar](60) NULL,[INSCRICAOESTADUAL] [varchar](20) NULL,[TELEFONE] [varchar](15) NULL,[FAX] [varchar](15) NULL,[EMAIL] [varchar](60) NULL,[RUA] [varchar](100) NULL,[NUMERO] [varchar](8) NULL,[COMPLEMENTO] [VARCHAR] (60) NULL,[BAIRRO] [VARCHAR] (80) NULL,[CIDADE] [VARCHAR] (32) NULL,[ESTADO] [varchar](2) NULL,[PAIS] [varchar](20) NULL,[CEP] [VARCHAR] (9) NULL, " & _
        "   [CONTROLACGC] [smallint] NULL,[CONTROLE1] [smallint] NULL,[CONTROLE2] [smallint] NULL,[CONTROLE3] [smallint] NULL,[IDIMAGEM] [int] NULL,[PRODUTORRURAL] [VARCHAR] (1) NULL,[ATIVO] [VARCHAR] (1) NULL,[CODEXTERNO] [varchar](10) NULL,[IMPORTADA] [VARCHAR] (1) NULL,[DATALIMITELICENCAS] [datetime] NULL) ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF " & _
        "ALTER TABLE [dbo].[GCOLIGADA] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[GCOLIGADA] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[GCOLIGADA] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[GCOLIGADA] ADD [RECMODIFIEDON] [datetime] NULL ALTER TABLE [dbo].[GCOLIGADA] ADD  CONSTRAINT [PKGCOLIGADA]  " & _
        "PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'PFUNC
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[PFUNC]( " & _
        "   [CODCOLIGADA] [SMALLINT] NOT NULL,[CHAPA] [VARCHAR] (16) NOT NULL,[NROFICHAREG] [int] NULL,[CODRECEBIMENTO] [char](1) NULL,[CODSITUACAO] [char](1) NULL,[CODTIPO] [char](1) NULL,[CODSECAO] [varchar](35) NULL,[CODFUNCAO] [varchar](10) NULL,[CODSINDICATO] [varchar](10) NULL,[JORNADA] [NUMERIC] (15,2) NULL,[CODHORARIO] [varchar](10) NULL,[NRODEPIRRF] [smallint] NULL,[NRODEPSALFAM] [smallint] NULL,[DTBASE] [datetime] NULL,[SALARIO] [NUMERIC] (15,2) NULL, " & _
        "   [SITUACAOFGTS] [char](1) NULL,[DTOPCAOFGTS] [datetime] NULL,[CONTAFGTS] [varchar](11) NULL,[SALDOFGTS] [NUMERIC] (15,2) NULL,[DTSALDOFGTS] [datetime] NULL,[CONTRIBSINDICAL] [char](1) NULL,[APOSENTADO] [smallint] NULL,[TEMMAIS65ANOS] [smallint] NULL,[AJUDACUSTO] [NUMERIC] (15,2) NULL,[PERCENTADIANT] [NUMERIC] (15,2) NULL,[ARREDONDAMENTO] [NUMERIC] (15,2) NULL,[DATAADMISSAO] [datetime] NULL,[TIPOADMISSAO] [char](1) NULL,[DTTRANSFERENCIA] [datetime] NULL,[MOTIVOADMISSAO] [varchar](2) NULL, " & _
        "   [TEMPRAZOCONTR] [smallint] NULL,[FIMPRAZOCONTR] [datetime] NULL,[DATADEMISSAO] [datetime] NULL,[TIPODEMISSAO] [char](1) NULL,[MOTIVODEMISSAO] [varchar](2) NULL,[DTDESLIGAMENTO] [datetime] NULL,[DTULTIMOMOVIM] [datetime] NULL,[DTPAGTORESCISAO] [datetime] NULL,[CODSAQUEFGTS] [varchar](2) NULL,[TEMAVISOPREVIO] [smallint] NULL,[DTAVISOPREVIO] [datetime] NULL,[NRODIASAVISO] [smallint] NULL,[DTVENCFERIAS] [datetime] NULL,[INICPROGFERIAS1] [datetime] NULL,[FIMPROGFERIAS1] [datetime] NULL, " & _
        "   [QUERABONO] [smallint] NULL,[QUER1APARC13O] [smallint] NULL,[NRODIASADIANTFER] [smallint] NULL,[EVTADIANTFERIAS] [varchar](4) NULL,[FERIASCOLETIVAS] [smallint] NULL,[NRODIASFERIAS] [NUMERIC] (15,2) NULL,[NRODIASABONO] [NUMERIC] (15,2) NULL,[INICPROGFERIAS2] [datetime] NULL,[FIMPROGFERIAS2] [datetime] NULL,[SALDOFERIAS] [NUMERIC] (15,2) NULL,[SALDOFERIASANT] [NUMERIC] (15,2) NULL,[SALDOFERANTAUX] [NUMERIC] (15,2) NULL,[OBSFERIAS] [varchar](80) NULL,[DTPAGTOFERIAS] [datetime] NULL,[DTAVISOFERIAS] [datetime] NULL, " & _
        "   [NDIASLICREM1] [NUMERIC] (15,2) NULL,[NDIASLICREM2] [NUMERIC] (15,2) NULL,[DTINICIOLICENCA] [datetime] NULL,[MEDIASALMATERN] [NUMERIC] (15,2) NULL,[SITUACAORAIS] [char](1) NULL,[CONTAPAGAMENTO] [varchar](15) NULL,[MEMBROSINDICAL] [smallint] NULL,[VINCULORAIS] [varchar](2) NULL,[USAVALETRANSP] [smallint] NULL,[DIASUTEISMES] [smallint] NULL,[DIASUTMEIOEXP] [smallint] NULL,[DIASUTPROXMES] [smallint] NULL,[DIASUTPROXMEIO] [smallint] NULL,[DIASUTRESTANTES] [smallint] NULL,[DIASUTRESTMEIO] [smallint] NULL, " & _
        "   [MUDOUENDERECO] [smallint] NULL,[MUDOUCARTTRAB] [smallint] NULL,[ANTIGACARTTRAB] [varchar](10) NULL,[ANTIGASERIECART] [varchar](5) NULL,[MUDOUNOME] [smallint] NULL,[ANTIGONOME] [varchar](120) NULL,[MUDOUPIS] [smallint] NULL,[ANTIGOPIS] [varchar](11) NULL,[MUDOUCHAPA] [smallint] NULL,[ANTIGACHAPA] [varchar](16) NULL,[MUDOUADMISSAO] [smallint] NULL,[ANTIGADTADM] [datetime] NULL,[ANTIGOVINCULO] [char](1) NULL,[ANTIGOTIPOFUNC] [char](1) NULL,[ANTIGOTIPOADM] [char](1) NULL, " & _
        "   [MUDOUDTOPCAO] [smallint] NULL,[ANTIGADTOPCAO] [datetime] NULL,[MUDOUSECAO] [smallint] NULL,[ANTIGASECAO] [varchar](35) NULL,[MUDOUDTNASCIM] [smallint] NULL,[ANTIGADTNASCIM] [datetime] NULL,[FALTAALTERFGTS] [smallint] NULL,[DEDUZIRRF65] [smallint] NULL,[PISPARAFGTS] [varchar](11) NULL,[ULTIMORECALCULODATA] [datetime] NULL,[ULTIMORECALCULOHORA] [datetime] NULL,[DESCONTAAVISOPREVIO] [smallint] NULL,[CODFILIAL] [smallint] NOT NULL,[NOME] [varchar](120) NULL,[INDINICIOHOR] [smallint] NULL, " & _
        "   [PISPASEP] [varchar](11) NULL,[DTCADASTROPIS] [datetime] NULL,[CODPESSOA] [int] NOT NULL,[CODBANCOFGTS] [varchar](3) NULL,[CODBANCOPAGTO] [varchar](3) NULL,[CODAGENCIAPAGTO] [varchar](6) NULL,[CODBANCOPIS] [varchar](3) NULL,[RESCISAOCALCULADA] [SMALLINT] NULL,[OPBANCARIA] [varchar](5) NULL,[MEMBROCIPA] [SMALLINT] NULL,[USASALCOMPOSTO] [SMALLINT] NULL,[REGATUAL] [int] NULL,[NUMVEZESDESCEMPRESTIMO] [smallint] NULL,[DATAINICIODESCEMPRESTIMO] [datetime] NULL,[GRUPOSALARIAL] [varchar](10) NULL,[JORNADAMENSAL] [smallint] NULL, " & _
        "   [PREVDISP] [datetime] NULL,[CODOCORRENCIA] [smallint] NOT NULL,[CODCATEGORIA] [smallint] NULL,[CLASSECONTRIB] [smallint] NULL,[CODEQUIPE] [varchar](20) NULL,[ESUPERVISOR] [SMALLINT] NULL,[INTEGRCONTABIL] [varchar](22) NULL,[INTEGRGERENCIAL] [varchar](22) NULL,[USACONTROLEDESALDO] [SMALLINT] NULL,[CI] [varchar](11) NULL,[MUDOUCI] [SMALLINT] NULL,[ANTIGOCI] [varchar](11) NULL,[PERIODORESCISAO] [smallint] NULL,[CODGRPQUIOSQUE] [varchar](15) NULL,[FGTSMESANTRECOLGRFP] [SMALLINT] NULL,[CODNIVELSAL] [varchar](10) NULL, " & _
        "   [TRABALHOUNADEMISSAO] [SMALLINT] NULL,[NRODIASFERIASJORNRED] [smallint] NULL,[POSSUIALVARAMENOR16] [smallint] NULL,[DATARESCISAO] [datetime] NULL,[SITUACAOINSS] [smallint] NULL,[DTAPOSENTADORIA] [datetime] NULL,[CODTABELASALARIAL] [varchar](10) NULL,[TEMDEDUCAOCPMF] [smallint] NULL,[NRODIASFERIASCORRIDOS] [smallint] NULL,[NRODIASABONOCORRIDOS] [smallint] NULL,[POSICAOABONO] [smallint] NULL,[REGIMEREVEZAMENTO] [varchar](15) NULL,[QUERADIANTAMENTO] [smallint] NULL,[DTPROXAQUISFERIAS] [datetime] NULL,[CODCOLFORNEC] [SMALLINT] NULL,[CODFORNECEDOR] [varchar](25) NULL, " & _
        "   [ISENTOIRRF] [smallint] NULL,[ANOCOMPTRANSF] [smallint] NULL,[MESCOMPTRANSF] [smallint] NULL,[NROPERIODOTRANSF] [smallint] NULL,[TIPOAPOSENTADORIA] [smallint] NULL,[REPOEVAGA] [varchar](1) NULL,[SALDOFGTSREAL] [NUMERIC] (15,2) NULL,[RESCISAOPRECISARECALC] [SMALLINT] NULL,[ID] [int] IDENTITY(1,1) NOT NULL,[BENEFPONTOS] [int] NULL) ON [PRIMARY] " & _
        "SET ANSI_PADDING OFF ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBASSOC1OCORRCNPJ] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBASSOC2OCORRCNPJ] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBASSISTCNPJ] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBCONFEDCNPJ] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBASSOC1OCORRVALOR] [NUMERIC] (15,2) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBASSOC2OCORRVALOR] [NUMERIC] (15,2) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBASSISTVALOR] [NUMERIC] (15,2) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTRIBCONFEDVALOR] [NUMERIC] (15,2) NULL ALTER TABLE [dbo].[PFUNC] ADD [LOCALTRABCODMUNCIPIO] [varchar](15) NULL ALTER TABLE [dbo].[PFUNC] ADD [MESESHORAEXTRAS] [int] NULL ALTER TABLE [dbo].[PFUNC] ADD [MESESGRATIFICACAO] [int] NULL ALTER TABLE [dbo].[PFUNC] ADD [MESESDISSIDIOCOLETIVO] [int] NULL " & _
        "ALTER TABLE [dbo].[PFUNC] ADD [INDICADORSINDICALIZADO] [SMALLINT] NULL ALTER TABLE [dbo].[PFUNC] ADD [DTULTIMOMOVIMPGTOEXTRAS] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [FIMPRAZOPRORROGCONTR] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [RECCREATEDBY] [varchar](50) NULL ALTER TABLE [dbo].[PFUNC] ADD [RECCREATEDON] [datetime] NULL ALTER TABLE [dbo].[PFUNC] ADD [RECMODIFIEDBY] [varchar](50) NULL ALTER TABLE [dbo].[PFUNC] ADD [RECMODIFIEDON] [datetime] NULL SET ANSI_PADDING ON ALTER TABLE [dbo].[PFUNC] ADD [NUMEROCARTAOSUS] [varchar](100) NULL ALTER TABLE [dbo].[PFUNC] ADD [CODCOLIGADAORIGEM] [SMALLINT] NULL ALTER TABLE [dbo].[PFUNC] ADD [CHAPAORIGEM] [VARCHAR] (16) NULL ALTER TABLE [dbo].[PFUNC] ADD [FERIASFINALIZADASPROXMES] [SMALLINT] NULL ALTER TABLE [dbo].[PFUNC] ADD [INDSIMPLES] [smallint] NULL " & _
        "SET ANSI_PADDING OFF ALTER TABLE [dbo].[PFUNC] ADD [TPCONTABANCARIA] [varchar](1) NULL ALTER TABLE [dbo].[PFUNC] ADD [NRPROCJUD] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [RESIDENCIAPROPRIA] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [RESIDENCIARECURSOSFGTS] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [TPREGIMEPREV] [varchar](1) NULL ALTER TABLE [dbo].[PFUNC] ADD [INDADMISSAO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [TIPOREINTEGRACAO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [DATAREINTEGRACAO] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [DATARETORNOEFETIVO] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [NROLEIANISTIA] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [NROPROCESSOJUDICIAL] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [NATUREZAESTAGIO] [char](1) NULL ALTER TABLE [dbo].[PFUNC] ADD [CODNIVELESTAGIO] [char](1) NULL " & _
        "ALTER TABLE [dbo].[PFUNC] ADD [AREAATUACAOESTAGIO] [varchar](50) NULL ALTER TABLE [dbo].[PFUNC] ADD [NUMEROAPOLICEESTAGIO] [varchar](30) NULL ALTER TABLE [dbo].[PFUNC] ADD [DTPREVTERMINOESTAGIO] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [CODINSTITUICAOENSINOESTAGIO] [varchar](16) NULL ALTER TABLE [dbo].[PFUNC] ADD [CODAGENTEINTEGRACAOESTAGIO] [varchar](16) NULL ALTER TABLE [dbo].[PFUNC] ADD [CPFCOORDENADORESTAGIO] [varchar](11) NULL ALTER TABLE [dbo].[PFUNC] ADD [NOMECOORDENADORESTAGIO] [varchar](80) NULL ALTER TABLE [dbo].[PFUNC] ADD [CNPJEMPRESAORIGEM] [varchar](14) NULL ALTER TABLE [dbo].[PFUNC] ADD [DTADMISSAOEMPRESAORIGEM] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [MATRICULAEMPRESAORIGEM] [varchar](30) NULL ALTER TABLE [dbo].[PFUNC] ADD [CODCATEGORIAEMPRESAORIGEM] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [TIPOREDUCAOAVISO] [smallint] NULL " & _
        "ALTER TABLE [dbo].[PFUNC] ADD [FORMAREDUCAOAVISO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [RECEBSEGDESEMP] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [MOTIVOTRABTEMP] [smallint] NULL SET ANSI_PADDING ON ALTER TABLE [dbo].[PFUNC] ADD [CHAPASUBSTRABTEMP] [varchar](16) NULL ALTER TABLE [dbo].[PFUNC] ADD [MOTIVOCANCELAMENTOAVISO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [DTCANCELAMENTOAVISO] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [NROATESTADOOBITO] [varchar](30) NULL ALTER TABLE [dbo].[PFUNC] ADD [NROPROCESSOTRAB] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [OBSERVACAORESCISAO] [varchar](255) NULL ALTER TABLE [dbo].[PFUNC] ADD [OBSERVACAOAVISOPREVIO] [varchar](255) NULL ALTER TABLE [dbo].[PFUNC] ADD [CARREGOUAVISOPREVIO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [DESCRICAOSALVARIAVEL] [varchar](90) NULL " & _
        "ALTER TABLE [dbo].[PFUNC] ADD [OBSCANCELAMENTOAVISO] [varchar](255) NULL ALTER TABLE [dbo].[PFUNC] ADD [SUCESSAOVINCULO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [CNPJEMPRESAANTERIOR] [varchar](20) NULL ALTER TABLE [dbo].[PFUNC] ADD [MATRICULAANTERIOR] [varchar](30) NULL ALTER TABLE [dbo].[PFUNC] ADD [DTINICIOVINCULO] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [OBSERVACAOSUCESSAO] [varchar](255) NULL ALTER TABLE [dbo].[PFUNC] ADD [TRANSFERENCIASUCESSAO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [APMISTO_DTAVTRAB] [datetime] NULL ALTER TABLE [dbo].[PFUNC] ADD [APMISTO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [FERIASDIASUTEIS] [SMALLINT] NULL ALTER TABLE [dbo].[PFUNC] ADD [FERIASSALDODIASUTEIS] [NUMERIC] (15,2) NULL ALTER TABLE [dbo].[PFUNC] ADD [SITUACAOIRRF] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [IDDADOSRESID] [smallint] NULL " & _
        "ALTER TABLE [dbo].[PFUNC] ADD [SEQUENCIATRANSF] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [CODBANCOPAGTO2] [varchar](3) NULL ALTER TABLE [dbo].[PFUNC] ADD [CODAGENCIAPAGTO2] [varchar](6) NULL ALTER TABLE [dbo].[PFUNC] ADD [CONTAPAGAMENTO2] [varchar](15) NULL ALTER TABLE [dbo].[PFUNC] ADD [OPBANCARIA2] [varchar](5) NULL ALTER TABLE [dbo].[PFUNC] ADD [TPCONTABANCARIA2] [varchar](1) NULL ALTER TABLE [dbo].[PFUNC] ADD [INDPAGTOJUIZO] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [CODIGORECEITA3533] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [DTDESLIGAMENTOREINT] [DATETIME] NULL ALTER TABLE [dbo].[PFUNC] ADD [MATRICULAESOCIAL] [varchar](30) NULL ALTER TABLE [dbo].[PFUNC] ADD [MOTIVOTRANSFERENCIA] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD [ANOSCONTRIBINSS] [int] NULL ALTER TABLE [dbo].[PFUNC] ADD [CODORGORIDES] [varchar](10) NULL " & _
        "ALTER TABLE [dbo].[PFUNC] ADD [CODREGJURI] [char](2) NULL ALTER TABLE [dbo].[PFUNC] ADD [CODCCUSTO] [varchar](25) NULL ALTER TABLE [dbo].[PFUNC] ADD [IDITEMCONTABIL] [int] NULL ALTER TABLE [dbo].[PFUNC] ADD [IDCLASSEVALOR] [int] NULL ALTER TABLE [dbo].[PFUNC] ADD [TIPOREGIMEJORNADA] [smallint] NULL ALTER TABLE [dbo].[PFUNC] ADD  CONSTRAINT [PKPFUNC] PRIMARY KEY CLUSTERED ([CODCOLIGADA] ASC,[CHAPA] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        
        'TITMMOVRELAC
        oConn.Execute "CREATE TABLE " & vBanco & ".[dbo].[TITMMOVRELAC]( " & _
        "   [IDMOVORIGEM] [INT] NOT NULL,[NSEQITMMOVORIGEM] [INT] NOT NULL,[CODCOLORIGEM] [SMALLINT] NOT NULL,[IDMOVDESTINO] [INT] NOT NULL,[NSEQITMMOVDESTINO] [INT] NOT NULL,[CODCOLDESTINO] [SMALLINT] NOT NULL,[QUANTIDADE] [NUMERIC] (15,4) NULL,[RECCREATEDBY] [varchar](50) NULL,[RECCREATEDON] [datetime] NULL,[RECMODIFIEDBY] [varchar](50) NULL,[RECMODIFIEDON] [datetime] NULL,[VALORRECEBIDO] [NUMERIC] (15,4) NULL, " & _
        "CONSTRAINT [PKTITMMOVRELAC] PRIMARY KEY CLUSTERED ([IDMOVORIGEM] ASC,[NSEQITMMOVORIGEM] ASC,[CODCOLORIGEM] ASC,[IDMOVDESTINO] ASC,[NSEQITMMOVDESTINO] ASC,[CODCOLDESTINO] ASC " & _
        ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
        
        CriarTabelasADOOFFLine = True
    
    End If
       
    oConn.Close
    Set oConn = Nothing
    Exit Function
Err1:
    'Msgbox "(ADO) Erro ao criar Tabela de dados: " & vbCrLf & Err.Number & " - Tabela já Existe - " & Err.Description, 16, "Mensagem de erro"
    CriarTabelasADOOFFLine = False
    Resume Next
    'Exit Function
End Function


Public Function filtroPadrao()
    Dim rsFiltroPadrao As New ADODB.Recordset
    Dim sqlFiltroPadrao As String

    sqlFiltroPadrao = "select a.nomefiltro,a.query,a.expressao,a.tipofiltro,a.usuario,a.modulo,case when b.idfiltro is null then 'N' else 'S' end padrao,a.idfiltro " & _
                      "from tbfiltro as a left join tbFiltroPadrao as b on a.idfiltro = b.idfiltro and '" & vLogin & "' = b.idinfo where a.modulo = '" & Formulario & "'   and b.idfiltro is not null"
    rsFiltroPadrao.Open sqlFiltroPadrao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFiltroPadrao.RecordCount > 0 Then
        FiltroGeral = rsFiltroPadrao.Fields(0)
        SqlLV = rsFiltroPadrao.Fields(1)
        vSubstituto = rsFiltroPadrao.Fields(2)
        vNovoFiltro = rsFiltroPadrao.Fields(2)
        vMantemExpressao = rsFiltroPadrao.Fields(2)
        vTituloFiltro = rsFiltroPadrao.Fields(0)
        vIdFiltro = rsFiltroPadrao.Fields(7)
    Else
        FiltroGeral = "Todos"
    End If
    rsFiltroPadrao.Close
    Set rsFiltroPadrao = Nothing

End Function

Public Function LocalString1(vQuery As String)
    Dim vContador As Integer
    
    If InStr(UCase(vNovoFiltro), UCase("like '[]'")) > 0 Then
        vContador = 1
        frmPassaParametro.Text1.Tag = "string"
        frmPassaParametro.Frame1.Caption = vTituloFiltro
        frmPassaParametro.lblCampo1.Caption = "Parâmetro:"
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("like '[]'")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, UCase("like '[]"), UCase("like ") & "'%" & vAlteraLike)
                vNovoFiltro = vSubstituto
                vContador = vContador + 1
            Else
                vSubstituto = Replace(vNovoFiltro, UCase("like '[]"), UCase("like ") & "'%" & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
    End If
    
    If InStr(UCase(vNovoFiltro), UCase("like '[datetime]'")) > 0 Then
        vContador = 1
        frmPassaParametro.Text1.Tag = "data"
        frmPassaParametro.Frame1.Caption = vTituloFiltro
        frmPassaParametro.Text1.Tag = "data"
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("like '[datetime]'")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "LIKE '[datetime]", UCase("='") & vAlteraLike)
                vNovoFiltro = vSubstituto
                vContador = vContador + 1
            Else
                vSubstituto = Replace(vNovoFiltro, "LIKE '[datetime]", UCase("='") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
'----------------
    End If
    If InStr(UCase(vNovoFiltro), UCase("BETWEEN")) > 0 Then
        vContador = 1
        frmPassaParametro.Frame1.Caption = vTituloFiltro
        frmPassaParametro.lblCampo1.Caption = "1ª data:"
        frmPassaParametro.lblCampo2.Caption = "2ª data:"
        
        frmPassaParametro.Text1.Tag = "data"
        frmPassaParametro.Text2.Tag = "data"
        frmPassaParametro.lblCampo2.Visible = True
        frmPassaParametro.Text2.Visible = True
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("'[datetime")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "[datetime1]", UCase("") & vAlteraLike)
                vContador = vContador + 1
            Else
                vNovoFiltro = vSubstituto
                vSubstituto = Replace(vNovoFiltro, "[datetime2]", UCase("") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
    End If
    If InStr(UCase(vNovoFiltro), UCase("'[datetime")) > 0 Then
        vContador = 1
        frmPassaParametro.Frame1.Caption = vTituloFiltro
        frmPassaParametro.Text1.Tag = "data"
        frmPassaParametro.lblCampo1.Caption = "Data:"
        If InStr(UCase(vNovoFiltro), UCase("'[datetime2")) > 0 Then
            frmPassaParametro.Frame1.Caption = vTituloFiltro
            frmPassaParametro.lblCampo1.Caption = "1ª data:"
            frmPassaParametro.lblCampo2.Caption = "2ª data:"
            frmPassaParametro.Text2.Tag = "data"
            frmPassaParametro.Text2.Visible = True
            frmPassaParametro.lblCampo2.Visible = True
        End If
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("'[datetime")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "[datetime1]", vAlteraLike)
                vContador = vContador + 1
            Else
                vNovoFiltro = vSubstituto
                vSubstituto = Replace(vNovoFiltro, "'[datetime2]", UCase("'") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
'-----------------
    End If
    If InStr(UCase(vNovoFiltro), UCase("IN([])")) > 0 Then
        vContador = 1
        frmPassaParametro.Text2.Tag = "string"
        frmPassaParametro.lblCampo1.Caption = "Informe os paramentros:"
        frmPassaParametro.Frame1.Caption = vTituloFiltro
        
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("IN([])")) > 0
            vAlteraLike = Replace(vAlteraLike, "%", "")
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "[]", UCase("") & vAlteraLike)
                vNovoFiltro = vSubstituto
                vContador = vContador + 1
            Else
                vNovoFiltro = vSubstituto
                vSubstituto = Replace(vNovoFiltro, "[]", UCase("") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
    End If
End Function

Public Function enviaEmail()
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
'DICA: CRIA O DOCUMENTO NO WORD, COPIA, ABRE O OUTLOOK, CRIE UM NOVO EMAIL, COLE E SALVE COMO HTML
'VC IRÁ TRABALHAR COM O ARQUIVO HTML GERADO.
'SE O ARQUIVO FOR MUITO GRANDE, CRIE VÁRIAS VARIÁVEIS DO TIPO STRING PARA ARMAZENAR O ARQUIVO PICADO
'QUANDO FOR ENVIAR CONCATENE TODAS AS VARIAVEIS
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
On Error GoTo errMail
    Dim HTMLBody1 As String, HTMLBody2 As String, HTMLBody3 As String, HTMLBody4 As String, HTMLBody5 As String, HTMLBody6 As String, HTMLBody7 As String
    Dim vCorDecisao As String
    Dim msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields

    enviaEmail = True


    With Camp
'PARAMETROS AUTOMATIZADOS
'        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
'        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
'        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
'        .Item(cdoSMTPServerPort) = vPorta 'Informe a porta SMTP
'        .Item(cdoSMTPAuthenticate) = vSMTPAutentic '1 '(MECANISMO DE AUTENTICAÇÃO)
'        .Item(cdoSMTPUseSSL).Value = vSSL '(INFORMA SE REQUER CONFIGURAÇÃO SEGURA)
'        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
'        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
'        .Update
        
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = "smtp.idg-eng.com"  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPServerPort) = 587 'Informe a porta SMTP
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSMTPUseSSL).Value = False
        .Item(cdoSendUserName) = "sistemas@idg-eng.com" ' informe o usuario de autenticação
        .Item(cdoSendPassword) = "gm990817"  'Informe a Senha aqui
        .Update
        
    End With

    With msg
        Set .Configuration = Cof
      
        .To = vEmailAprovadores
        .From = "sistemas@idg-eng.com"
        .Subject = "BLOQUEIO DE MEDICAO DE SUBCONTRATADOS"
        .CC = sEmailAvRec 'Emails em copia
       
        HTMLBody1 = "<html>" & _
        "<head> " & _
        "<meta http-equiv=Content-Type content='text/html; charset=windows-1252'>'" & _
        "<meta name=ProgId content=Word.Document>" & _
        "<meta name=Generator content='Microsoft Word 14'>" & _
        "<meta name=Originator content='Microsoft Word 14'>" & _
        "<link rel=File-List" & _
        "</head>" & _
        "<body lang=PT-BR link=blue vlink=purple style='tab-interval:35.4pt'>" & _
        "<p>Informamos que a medi&ccedil;&atilde;o n&uacute;mero: <span style='color:#FF0000'> " & vNumMedicao & "</span>, do colaborador: <span style='color:#FF0000'> " & vColaborador & "</span>, " & _
        "referente ao per&iacute;odo:<span style='color:#FF0000'> " & vPeriodoMedicao & ",</span> compet&ecirc;ncia <span style='color:#FF0000'>" & vCompetenciaMedicao & "</span> foi bloqueada pelo " & _
        "setor fiscal da IDG Engenharia e Consultoria e automaticamente rejeitada no sistema apontamento.info da IDG, pelo motivo apresentado a seguir:</p> " & _
        "<p><span style='color:#FF0000'>" & vTextMotivoBloqueio & "</span></p> " & _
        "<p>Entrar em contato com o setor fiscal para maiores esclarecimentos.</p> " & _
        "<p>Atenciosamente,</p> " & _
        "<p><p>Setor Fiscal</p> " & _
        "<h2 style='font-style:italic'>&nbsp;</h2>" & _
        "</body> " & _
        "</html> "
       
        .HTMLBody = HTMLBody1
        .Send
    End With
    Exit Function
errMail:
    enviaEmail = False
    If Err.Number = -2147220977 Then
'        mobjMsg.Abrir "O endereço foi rejeitado pelo servidor de email " & vEmailAprovadores & " " & sEmailAvRec, Ok, critico
        MsgBox "O endereço foi rejeitado pelo servidor de email " & vEmailAprovadores & " " & sEmailAvRec
        
        Resume Next
        Exit Function
    End If
    MsgBox "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SAF"
    Exit Function
End Function

'Public Function MarcaDesmarca(LV As ListView)
'    'Adiciona processo ao item selecionado no Listview
'    Dim Y As Integer, X As Integer
'
'    Y = LV.ListItems.Count
'    For X = 1 To Y
'        LV.ListItems(X).Selected = True
'        If LV.ListItems.Item(X).Checked = True Then
'            LV.ListItems.Item(X).Checked = False
'        Else
'            LV.ListItems.Item(X).Checked = True
'        End If
'    Next
'End Function

