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

Public mobjMsg As Msgbox

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

Public Sub ChangeRes(iWidth As Single, iHeight As Single)
   Dim A As Boolean
   Dim i As Long
   Do
      A = EnumDisplaySettings(0&, i, DevM)
      i = i + 1
   Loop Until (A = False)

   Dim b As Long
   DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
   DevM.dmPelsWidth = iWidth
   DevM.dmPelsHeight = iHeight
   b = ChangeDisplaySettings(DevM, 0)
End Sub

Public Function Img()
CaMinho = Servidor & ":"
Set Principal.Image1.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")
End Function

Public Function AplicarSkin(Frm As Form, Skin As Skin)
    CaminhoSkin = App.Path & "\MySkin.Skn"
    Skin.LoadSkin CaminhoSkin
    Skin.ApplySkin Frm.hwnd
    
    Set mobjMsg = New Msgbox
    'mobjMsg.Skin App.Path & "\MySkin.skn"
End Function

'funcao para ler valor de Chave
'Public Function iniReadKey(FileName As String, section As String, Key As String) As String
'    Dim RetVal As String * 255, v As Long
'    v = GetPrivateProfileString(section, Key, "", RetVal, 255, FileName)
'    iniReadKey = Left(RetVal, v)
'End Function

'Public Sub TemaMenu()
'    Tema = (iniReadKey(App.Path & "\config.ini", "TEMA", "NomeTema"))
'End Sub

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
            Msgbox "grava_Dados Nº do erro: " & Err.Number & ", na linha: " & Str$(Erl) & vbCrLf & _
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

Public Function ClonarDados(vServer As String, vBanco As String, vUsuario As String, vSenha As String, vQualTabela As String, vCondicao As String)
On Error GoTo Err
    Dim rsTabelaTotvs As New ADODB.Recordset
    Dim sqlTabelaTotvs As String
    Dim rsTabelaOFF As New ADODB.Recordset
    Dim sqlTabelaOFF As String
     
    Dim rsTamanhoTabela As New ADODB.Recordset
    Dim sqlTamanhoTabela As String
    Dim vQtdColsTabela As Integer, X As Double
     
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenha & ";Persist Security Info=True;User ID=" & vUsuario & ";Initial Catalog=" & vBanco & ";Data Source=" & vServer
     
    If vBanco = "FERRAMENTARIA_OFF" Then
        sqlTamanhoTabela = "select count(*) from information_schema.columns Where Table_Name='" & vQualTabela & "'"
        rsTamanhoTabela.Open sqlTamanhoTabela, cnBanco, adOpenKeyset, adLockReadOnly
        vQtdColsTabela = rsTamanhoTabela.Fields(0)
    
        sqlTabelaTotvs = "Select * from " & vQualTabela & " " & vCondicao
        rsTabelaTotvs.Open sqlTabelaTotvs, cnBanco, adOpenKeyset, adLockReadOnly
    ElseIf vBanco = "CORPORERM_OFF" Then
        sqlTamanhoTabela = "select count(*) from information_schema.columns Where Table_Name='" & vQualTabela & "'"
        rsTamanhoTabela.Open sqlTamanhoTabela, cnBancoSAP, adOpenKeyset, adLockReadOnly
        vQtdColsTabela = rsTamanhoTabela.Fields(0)
    
        sqlTabelaTotvs = "Select * from " & vQualTabela & " " & vCondicao
        rsTabelaTotvs.Open sqlTabelaTotvs, cnBancoSAP, adOpenKeyset, adLockReadOnly
    End If
    
    Legenda = "Aguarde, selecionando dados da tabela '" & vQualTabela
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    Principal.ProgressBar1.Max = rsTabelaTotvs.RecordCount

    Legenda = "Aguarde, importando tabela '" & vQualTabela '& "' Registro: " & rsTabelaTotvs.AbsolutePosition & ""
    Principal.StatusBar1.Panels(3).Text = Legenda
    Dim Y As Double
    Y = 0
    sqlTabelaOFF = "Select * from " & vQualTabela
    rsTabelaOFF.Open sqlTabelaOFF, oConn, adOpenKeyset, adLockOptimistic
    While Not rsTabelaTotvs.EOF
       rsTabelaOFF.AddNew
        For X = 0 To vQtdColsTabela - 1
            rsTabelaOFF.Fields(X) = rsTabelaTotvs.Fields(X)
        Next
        rsTabelaTotvs.MoveNext
        rsTabelaOFF.Update
        Y = Y + 1
        Principal.ProgressBar1.Value = Y
    Wend
    Set rsTabelaOFF = Nothing
    Principal.ProgressBar1.Value = 0
    rsTabelaTotvs.Close
    Set rsTabelaTotvs = Nothing
    
    rsTamanhoTabela.Close
    Set rsTamanhoTabela = Nothing
    Exit Function
Err:
    If Err.Number = 3709 Then
        ConexaoSAP
        rsTamanhoTabela.Open sqlTamanhoTabela, cnBancoSAP, adOpenKeyset, adLockReadOnly
    End If
    Resume Next
End Function

Public Function ControlaRegTabs(nomeTabelaP As String, nomeTabelaS As String, nomeCampo As String, nomeCampo2 As String, vBanco As String)
On Error Resume Next
    Dim vDataTime1 As String
    Dim vDataTime2 As String
'APENAS PARA TESTE
    vServerOffline = frmConfSistema.txtIntegra(13).Text 'Servidor e porta
    If vBanco = "FERRAMENTARIA" Then
        vBancoOffline = frmConfSistema.txtIntegra(9).Text 'Banco
    Else
        vBancoOffline = frmConfSistema.txtIntegra(12).Text 'Banco
    End If
    vUsuBancoOffline = frmConfSistema.txtIntegra(11).Text 'Usuário
    vSenhaBancoOffline = frmConfSistema.txtIntegra(10).Text 'Senha
   
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenhaBancoOffline & ";Persist Security Info=True;User ID=" & vUsuBancoOffline & ";Initial Catalog=" & vBancoOffline & ";Data Source=" & vServerOffline
'APENAS PARA TESTE
    
    Dim sqlTabOFF As String
    Dim rsTabOFF As New ADODB.Recordset
    
    Dim sqlControlTabs As String
    Dim rsControlTabs As New ADODB.Recordset
    If vBanco <> "FERRAMENTARIA" Then
    
        If nomeCampo2 = "" Then
            sqlTabOFF = "Select top 1 a." & nomeCampo & " from " & nomeTabelaP & " as a order by a." & nomeCampo & " desc"
        Else
            sqlTabOFF = "Select top 1 a." & nomeCampo & ",a." & nomeCampo2 & " from " & nomeTabelaP & " as a order by a." & nomeCampo & " desc, " & nomeCampo2 & " desc"
        End If
        rsTabOFF.Open sqlTabOFF, oConn, adOpenKeyset, adLockReadOnly
    
        vDataTime1 = Format(rsTabOFF.Fields(0), "mm/dd/yyyy hh:mm:ss")
    
        If nomeCampo2 = "" Then
            sqlControlTabs = "Insert into tbControlInsertRec(nometabelaP,nomeTabelaS,datacontrole1) Values('" & nomeTabelaP & "','" & nomeTabelaS & "',convert(varchar ,'" & vDataTime1 & "',121))"
        Else
            sqlControlTabs = "Insert into tbControlInsertRec(nometabelaP,nomeTabelaS,datacontrole1,datacontrole2) Values('" & nomeTabelaP & "','" & nomeTabelaS & "',convert(varchar ,'" & vDataTime1 & "',121))"
        End If
        rsControlTabs.Open sqlControlTabs, cnBanco
    Else
'    ControlaRegTabs "TBUSUARIOS", "-", "CODIGO", "", "FERRAMENTARIA"
        Dim vCodTBUsuario As String
        sqlTabOFF = "Select top 1 a." & nomeCampo & " from " & nomeTabelaP & " as a order by a." & nomeCampo & " desc"
        rsTabOFF.Open sqlTabOFF, oConn, adOpenKeyset, adLockReadOnly
        vCodTBUsuario = Format(rsTabOFF.Fields(0), "000000")
         vDataTime1 = Format(Date, "mm/dd/yyyy hh:mm:ss")
        sqlControlTabs = "Insert into tbControlInsertRec(nometabelaP,nomeTabelaS,datacontrole1) Values('" & nomeTabelaP & "','" & vCodTBUsuario & "',convert(varchar ,'" & vDataTime1 & "',121))"
        rsControlTabs.Open sqlControlTabs, cnBanco
    End If
    
    rsTabOFF.Close
    Set rsTabOFF = Nothing
End Function

Public Function ClonarDadosSvRemoto(vServer As String, vBanco As String, vUsuario As String, vSenha As String, vQualTabela As String, vCondicao As String, vBancoOFF As String, vNomeCampo As String)
On Error GoTo Err
    Dim rsTabelaTotvs As New ADODB.Recordset
    Dim sqlTabelaTotvs As String
    Dim rsTabelaOFF As New ADODB.Recordset
    Dim sqlTabelaOFF As String
     
    Dim rsControlTabs As New ADODB.Recordset
    Dim sqlControlTabs As String
     
    Dim rsTamanhoTabela As New ADODB.Recordset
    Dim sqlTamanhoTabela As String
    Dim vQtdColsTabela As Integer, X As Double
     
    Dim vUltimaDataHora As String
    Dim vCodTbUsuarios As Integer
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenha & ";Persist Security Info=True;User ID=" & vUsuario & ";Initial Catalog=" & vBanco & ";Data Source=" & vServer
    
    sqlTabelaTotvs = "Select * from tbcontrolInsertRec as a where a.nometabelaP = '" & vQualTabela & "'"
    rsTabelaTotvs.Open sqlTabelaTotvs, cnBanco, adOpenKeyset, adLockReadOnly
    vUltimaDataHora = Format(rsTabelaTotvs.Fields(2), "mm/dd/yyyy hh:mm:ss")
    If vBancoOFF = "FERRAMENTARIA_OFF" Then
        vCodTbUsuarios = Val(rsTabelaTotvs.Fields(1))
    End If
    
    rsTabelaTotvs.Close
    Set rsTabelaTotvs = Nothing
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenha & ";Persist Security Info=True;User ID=" & vUsuario & ";Initial Catalog=" & vBanco & ";Data Source=" & vServer
     
    If vBancoOFF = "CORPORERM_OFF" Then
        sqlTamanhoTabela = "select count(*) from information_schema.columns Where Table_Name='" & vQualTabela & "'"
        rsTamanhoTabela.Open sqlTamanhoTabela, oConn, adOpenKeyset, adLockReadOnly
        vQtdColsTabela = rsTamanhoTabela.Fields(0)
    
        If vQualTabela <> "TVEN" Then
            sqlTabelaTotvs = "Select * from " & vQualTabela & " as a where a." & vNomeCampo & " >  convert(varchar ,'" & vUltimaDataHora & "',121) order by a." & vNomeCampo
        Else
            sqlTabelaTotvs = "select a.* from " & vQualTabela & " as a inner join PFUNC as b on a.CODVEN = b.CHAPA where b." & vNomeCampo & " >  convert(varchar ,'" & vUltimaDataHora & "',121) order by b." & vNomeCampo
        End If
        rsTabelaTotvs.Open sqlTabelaTotvs, oConn, adOpenKeyset, adLockReadOnly
    Else
        sqlTamanhoTabela = "select count(*) from information_schema.columns Where Table_Name='" & vQualTabela & "'"
        rsTamanhoTabela.Open sqlTamanhoTabela, oConn, adOpenKeyset, adLockReadOnly
        vQtdColsTabela = rsTamanhoTabela.Fields(0)
    
        sqlTabelaTotvs = "Select * from " & vQualTabela & " as a where a." & vNomeCampo & " >  " & vCodTbUsuarios & " order by a." & vNomeCampo
        rsTabelaTotvs.Open sqlTabelaTotvs, oConn, adOpenKeyset, adLockReadOnly
    End If
    
    Legenda = "Aguarde, selecionando dados da tabela '" & vQualTabela
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    If rsTabelaTotvs.RecordCount > 0 Then Principal.ProgressBar1.Max = rsTabelaTotvs.RecordCount

    Legenda = "Aguarde, importando tabela '" & vQualTabela '& "' Registro: " & rsTabelaTotvs.AbsolutePosition & ""
    Principal.StatusBar1.Panels(3).Text = Legenda
    Dim Y As Double
    Y = 0
    sqlTabelaOFF = "Select * from " & vQualTabela
    If vBancoOFF = "CORPORERM_OFF" Then
        rsTabelaOFF.Open sqlTabelaOFF, cnBancoSAP, adOpenKeyset, adLockOptimistic
    Else
        rsTabelaOFF.Open sqlTabelaOFF, cnBanco, adOpenKeyset, adLockOptimistic
    End If
    While Not rsTabelaTotvs.EOF
       rsTabelaOFF.AddNew
        For X = 0 To vQtdColsTabela - 1
            rsTabelaOFF.Fields(X) = rsTabelaTotvs.Fields(X)
        Next
        rsTabelaTotvs.MoveNext
        rsTabelaOFF.Update
        Y = Y + 1
        Principal.ProgressBar1.Value = Y
    Wend
    Set rsTabelaOFF = Nothing
    Principal.ProgressBar1.Value = 0
    rsTabelaTotvs.Close
    Set rsTabelaTotvs = Nothing
    
    rsTamanhoTabela.Close
    Set rsTamanhoTabela = Nothing
    
    If vBancoOFF = "FERRAMENTARIA_OFF" Then
        sqlControlTabs = "update tbControlInsertRec set nometabelaS = (select right('00000' + rtrim(max(codigo)),6) from " & vQualTabela & ") where nometabelaP = '" & vQualTabela & "'"
        rsControlTabs.Open sqlControlTabs, cnBanco
    Else
        sqlTabelaTotvs = "SELECT MAX(" & vNomeCampo & ") FROM " & vQualTabela & ""
        rsTabelaTotvs.Open sqlTabelaTotvs, oConn, adOpenKeyset, adLockReadOnly
        vUltimaDataHora = Format(rsTabelaTotvs.Fields(0), "mm/dd/yyyy hh:mm:ss")
        rsTabelaTotvs.Close
        Set rsTabelaTotvs = Nothing
        
        sqlControlTabs = "update FERRAMENTARIA_OFF.dbo.tbControlInsertRec set datacontrole1 = (SELECT MAX(RECCREATEDON) FROM " & vQualTabela & ") where FERRAMENTARIA_OFF.dbo.tbControlInsertRec.nometabelaP = '" & vQualTabela & "'"
        rsControlTabs.Open sqlControlTabs, cnBancoSAP
    End If
    
    Exit Function
Err:
    If Err.Number = 3709 Then
        ConexaoSAP
        rsTamanhoTabela.Open sqlTamanhoTabela, cnBancoSAP, adOpenKeyset, adLockReadOnly
    End If
    Resume Next
End Function


'Rotina em estudo
Public Function SincronizarDadosImportar(vServerRemoto As String, vBancoRemoto As String, vUsuarioRemoto As String, vSenhaRemoto As String, vQualTabela As String, vCondicao As String, vOrdem As String)
    'vServerRemoto -> Nome do SERVIDOR que será conectado remotamente
    'vBancoRemoto -> Nome do BANCO que será conectado remotamente
    'vUsuarioRemoto -> Usuário de conexão ao BANCO que será conectado remotamente
    'vSenhaRemoto -> Senha de conexão ao BANCO que será conectado remotamente
    'vQualTabela -> Nome da tabela a ser trabalhada que se encontra no SERVIDOR remoto
    'vCondicao -> Condição WHERE da string de conexão
    
    Dim rsTabelaRemota As New ADODB.Recordset
    Dim sqlTabelaRemota As String
    Dim rsTabelaOFF As New ADODB.Recordset
    Dim sqlTabelaOFF As String
     
    Dim rsTamanhoTabela As New ADODB.Recordset
    Dim sqlTamanhoTabela As String
    
    Dim vQtdColsTabelaRemota As Integer, X As Double
    Dim vQtdRegTabelaRemota As Integer
    Dim Y As Double
     
    Legenda = "Aguarde, CONECTANDO AO SERVIDOR " & vServerRemoto
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenhaRemoto & ";Persist Security Info=True;User ID=" & vUsuarioRemoto & ";Initial Catalog=" & vBancoRemoto & ";Data Source=" & vServerRemoto

    'Captura a quantidade de colunas que a tabela REMOTA possui
    sqlTamanhoTabela = "select count(*) from information_schema.columns Where Table_Name='" & vQualTabela & "'"
    rsTamanhoTabela.Open sqlTamanhoTabela, oConn, adOpenKeyset, adLockReadOnly
    vQtdColsTabelaRemota = rsTamanhoTabela.Fields(0)
    rsTamanhoTabela.Close
    Set rsTamanhoTabela = Nothing
    
    'Saber a quantidade de registros da tabela REMOTA
    sqlTamanhoTabela = "select COUNT(*) as registros from " & vQualTabela & " " & vCondicao
    rsTamanhoTabela.Open sqlTamanhoTabela, oConn, adOpenKeyset, adLockReadOnly
    vQtdRegTabelaRemota = rsTamanhoTabela.Fields(0)
    
    'Da um select na tabela REMOTA
    sqlTabelaRemota = "Select * from " & vQualTabela & " " & vCondicao & " " & vOrdem
    rsTabelaRemota.Open sqlTabelaRemota, oConn, adOpenKeyset, adLockReadOnly
    Dim vEstoqueProdutoOFFLINE As Integer
    
    
    Legenda = "Aguarde, sincronizando dados da tabela '" & vQualTabela
    Principal.StatusBar1.Panels(3).Text = Legenda
    Principal.ProgressBar1.Max = rsTabelaRemota.RecordCount
    Y = 0
    
    While Not rsTabelaRemota.EOF
        'Da um select na tabela OFFLINE e verificando se o produto existe
        sqlTabelaOFF = "Select * from " & vQualTabela & " " & vCondicao & " and idprd = '" & rsTabelaRemota.Fields(3) & "' " & vOrdem
        rsTabelaOFF.Open sqlTabelaOFF, cnBancoSAP, adOpenKeyset, adLockOptimistic
        '(UPDATE)Se encontrar faça a rotina abaixo
        If rsTabelaOFF.RecordCount > 0 Then
            Dim rsEmpItens As New ADODB.Recordset
            Dim sqlEmpItens As String
            
            'Acha quanto do produto esta emprestado na tabela OFFLINE
            sqlEmpItens = "SELECT SUM(A.QTDEMPRESTADO-QTDDEVOLVIDA) AS TOTALEMPRESTADO FROM tbEmprestimoItens AS A WHERE A.idprd = '" & rsTabelaRemota.Fields(3) & "' GROUP BY idprd " & vOrdem
            rsEmpItens.Open sqlEmpItens, cnBanco, adOpenKeyset, adLockReadOnly
            'Estoque OFFLINE abaixo
            If rsEmpItens.RecordCount > 0 Then
                vEstoqueProdutoOFFLINE = rsTabelaOFF.Fields(5) + rsEmpItens.Fields(0)
            Else
                vEstoqueProdutoOFFLINE = rsTabelaOFF.Fields(5)
            End If
            
            If rsTabelaRemota.Fields(5) > vEstoqueProdutoOFFLINE Then
                rsTabelaOFF.Fields(5) = rsTabelaOFF.Fields(5) + (rsTabelaRemota.Fields(5) - vEstoqueProdutoOFFLINE)
                rsTabelaOFF.Fields(15) = rsTabelaOFF.Fields(5) * rsTabelaOFF.Fields(24)
            End If
            rsTabelaOFF.Update
            rsEmpItens.Close
            Set rsEmpItens = Nothing
        '(INSERT) Se Não encontrar faça a rotina abaixo
        Else
            rsTabelaOFF.AddNew
            For X = 0 To vQtdColsTabelaRemota - 1
                rsTabelaOFF.Fields(X) = rsTabelaRemota.Fields(X)
            Next
            rsTabelaOFF.Update
        End If
        rsTabelaOFF.Close
        Set rsTabelaOFF = Nothing
        rsTabelaRemota.MoveNext
        
        Y = Y + 1
        Principal.ProgressBar1.Value = Y
    Wend
    rsTabelaRemota.Close
    Set rsTabelaRemota = Nothing
End Function

Public Function SincronizarDadosExportar(vServerRemoto As String, vBancoRemoto As String, vUsuarioRemoto As String, vSenhaRemoto As String, vQualTabela As String, vCondicao As String, vOrdem As String, vQualDBSincronizar As String)
On Error GoTo Err
    'SINCRONIZA DADOS DO SERVIDOR OFFLINE (FERRAMENTARIA_OFF e CORPORERM_OFF) E
    'COM O SERVIDOR REMOTO (GNV e FERRAMENTARIA)
    
    'vServerRemoto -> Nome do SERVIDOR que será conectado remotamente
    'vBancoRemoto -> Nome do BANCO que será conectado remotamente
    'vUsuarioRemoto -> Usuário de conexão ao BANCO que será conectado remotamente
    'vSenhaRemoto -> Senha de conexão ao BANCO que será conectado remotamente
    'vQualTabela -> Nome da tabela a ser trabalhada que se encontra no SERVIDOR remoto
    'vCondicao -> Condição WHERE da string de conexão

    Dim rsTabelaRemota As New ADODB.Recordset
    Dim sqlTabelaRemota As String
    Dim rsTabelaOFF As New ADODB.Recordset
    Dim sqlTabelaOFF As String
    Dim rsSincronizar As New ADODB.Recordset
    Dim sqlSincronizar As String

    Dim rsTamanhoTabela As New ADODB.Recordset
    Dim sqlTamanhoTabela As String

    Dim vQtdColsTabelaRemota As Integer, X As Double
    Dim vQtdRegTabelaRemota As Integer


    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenhaRemoto & ";Persist Security Info=True;User ID=" & vUsuarioRemoto & ";Initial Catalog=" & vBancoRemoto & ";Data Source=" & vServerRemoto

    'Identifica quais movimentos serão sincronizados
    
'    If vQualTabela = "tbSincronizacao" Then
'        sqlSincronizar = "SELECT a.idmov FROM (SELECT idmov FROM tbEmprestimo where codcoligada = '" & vCodColigadaRM & "' UNION ALL SELECT idmov FROM tbDevolucao where codcoligada = '" & vCodColigadaRM & "') as A LEFT JOIN tbSincronizacao AS B ON a.idmov = B.idmovsincronizado WHERE B.idmovsincronizado IS NOT NULL ORDER BY A.idmov"
'    Else
        sqlSincronizar = "SELECT a.idmov FROM (SELECT idmov FROM tbEmprestimo where codcoligada = '" & vCodColigadaRM & "' UNION ALL SELECT idmov FROM tbDevolucao where codcoligada = '" & vCodColigadaRM & "') as A LEFT JOIN tbSincronizacao AS B ON a.idmov = B.idmovsincronizado WHERE B.idmovsincronizado IS NULL ORDER BY A.idmov"
'    End If
    rsSincronizar.Open sqlSincronizar, cnBanco, adOpenKeyset, adLockReadOnly
    
    
    'Captura a quantidade de colunas que a tabela REMOTA possui
    sqlTamanhoTabela = "select count(*) from information_schema.columns Where Table_Name='" & vQualTabela & "'"
    rsTamanhoTabela.Open sqlTamanhoTabela, oConn, adOpenKeyset, adLockReadOnly
    vQtdColsTabelaRemota = rsTamanhoTabela.Fields(0)
    rsTamanhoTabela.Close
    Set rsTamanhoTabela = Nothing
    
    While Not rsSincronizar.EOF
        'Da um select na tabela OFF de acordo com o posicionamento do IDMOV na tabela TBSINCRONIZAÇÃO
        sqlTabelaOFF = "Select * from " & vQualTabela & " " & vCondicao & "'" & rsSincronizar.Fields(0) & "'"
        If vQualDBSincronizar = "TOTVS" Then
            rsTabelaOFF.Open sqlTabelaOFF, cnBancoSAP, adOpenKeyset, adLockReadOnly
        Else
            rsTabelaOFF.Open sqlTabelaOFF, cnBanco, adOpenKeyset, adLockReadOnly
        End If
        
        
        sqlTabelaRemota = "SELECT TOP 1 * FROM " & vQualTabela & ""
        rsTabelaRemota.Open sqlTabelaRemota, oConn, adOpenKeyset, adLockOptimistic
        
        'INICIA A SINCRONIZAÇÃO DO IDMOV (PODE TER 1 OU MAIS REGISTROS - DEPENDE DA TABELA)
        'On Error Resume Next
        While Not rsTabelaOFF.EOF
            rsTabelaRemota.AddNew
            For X = 0 To vQtdColsTabelaRemota - 1
                rsTabelaRemota.Fields(X) = rsTabelaOFF.Fields(X)
            Next
            rsTabelaOFF.MoveNext
        Wend
        rsTabelaRemota.Update
        rsTabelaRemota.Close
        Set rsTabelaRemota = Nothing
        
        rsTabelaOFF.Close
        Set rsTabelaOFF = Nothing
        'TERMINA A SINCRONIZAÇÃO DO IDMOV
        rsSincronizar.MoveNext
    Wend
    rsSincronizar.Close
    Set rsSincronizar = Nothing
    
    oConn.Close
    Set oConn = Nothing
    
    
    If vQualTabela = "tbSincronizacao" Then
        On Error GoTo Err
        
        Dim rsAcertaTMOV As New ADODB.Recordset
        Dim sqlAcertaTMOV As String

        
        Set oConn = New ADODB.Connection
        oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenhaRemoto & ";Persist Security Info=True;User ID=" & vUsuarioRemoto & ";Initial Catalog=" & frmConfSistema.txtIntegra(12).Text & ";Data Source=" & vServerRemoto
        
        'sqlTabelaOFF = "Select max(idmovsincronizado)+1 as idmovsincronizado from " & vQualTabela & " "
        'rsTabelaOFF.Open sqlTabelaOFF, cnBanco, adOpenKeyset, adLockReadOnly
        
        'Atualiza GAUTOINC Local
        sqlAcertaTMOV = "UPDATE GAUTOINC set VALAUTOINC = " & vMaiorIDMOV & " where codautoinc like 'IDMOV' and codcoligada = '" & vCodColigadaRM & "'"
        rsAcertaTMOV.Open sqlAcertaTMOV, cnBancoSAP
    
        'Atualiza GAUTOINC Remoto
        sqlAcertaTMOV = "UPDATE GAUTOINC set VALAUTOINC = " & vMaiorIDMOV & " where codautoinc like 'IDMOV' and codcoligada = '" & vCodColigadaRM & "'"
        rsAcertaTMOV.Open sqlAcertaTMOV, oConn
        
        rsTabelaOFF.Close
        Set rsTabelaOFF = Nothing
    
        oConn.Close
        Set oConn = Nothing
    End If
    Exit Function
Err:
    'If Err.Number = -2147217873 Then mobjMsg.Abrir "Dados já foram importados em outro momento", Ok, critico, "Ferramentaria"
    'mobjMsg.Abrir "Ocorreu um erro: " & Err.Number & " " & Err.Description, Ok, critico, "Ferramentaria"
    Resume Next
End Function

Public Function ImportarDadosSistemaAntigo()
On Error GoTo Err
    Dim rsBuscaDadosAntigos As New ADODB.Recordset
    Dim sqlBuscaDadosAntigos As String
    
    Dim rsGravaEmprestimo As New ADODB.Recordset
    Dim sqlGravaEmprestimo As String
    
    Dim rsGravaEmprestimoItens As New ADODB.Recordset
    Dim sqlGravaEmprestimoItens As String
    
    Dim rsCompletaDados As New ADODB.Recordset
    Dim sqlCompletaDados As String
    Dim vNome As String, vCodFuncao As String, vNomeFuncao As String, vCodSecao As String, vNomeSecao As String
    Dim vIDMovLocal As Double
    
    sqlBuscaDadosAntigos = "SELECT codloc = M.CODLOC,L.NOME LOCAL,M.CODVEN1 as FUNCIONARIO,DATAEMISSAO = M.DATAEMISSAO,dife = CONVERT(VARCHAR, DATEDIFF(DAY, M.DATAEMISSAO ,GETDATE()) ),qtDiasEmp = p.CAMPOLIVRE,atrasoDev =  p.CAMPOLIVRE - CONVERT(VARCHAR, DATEDIFF(DAY, M.DATAEMISSAO ,GETDATE()) ),recolher  =  case when (p.CAMPOLIVRE - CONVERT(VARCHAR, DATEDIFF(DAY, M.DATAEMISSAO ,GETDATE()) )) <0 then  'Sim' else 'Não' end,manutencao = ( case when (SELECT manu.DATAVENCIMENTO from OFVENCPLANOMANUT manu INNER join TPRODUTO Prd on manu.IDOBJOF = SUBSTRING(Prd.CODIGOPRD,4,9) AND PRD.CODIGOPRD  = P.CODIGOPRD) < GETDATE() then 'Sim' else 'Não' end), " & _
                           "M.HORULTIMAALTERACAO,M.NUMEROMOV,M.SERIE,M.SEQUENCIALESTOQUE,CODIGOPRD  =P.CODIGOPRD,NOMEFANTASIA =P.NOMEFANTASIA,'-' AS CODBEM,I.IDMOV,I.IDPRD,I.NSEQITMMOV,I.QUANTIDADE,ISNULL((SELECT SUM(I2.QUANTIDADE) FROM TITMMOVRELAC(NOLOCK),TITMMOV I2,TMOV M2 WHERE M2.IDMOV=I2.IDMOV AND TITMMOVRELAC.IDMOVORIGEM=I2.IDMOV AND TITMMOVRELAC.NSEQITMMOVORIGEM=I2.NSEQITMMOV AND TITMMOVRELAC.IDMOVDESTINO = M.IDMOV AND TITMMOVRELAC.NSEQITMMOVDESTINO = I.NSEQITMMOV AND M2.CODTMV='1.2.16' ),0)AS QTDEDEVOLVIDA,I.QUANTIDADE - ISNULL((SELECT SUM(I2.QUANTIDADE) FROM TITMMOVRELAC(NOLOCK),TITMMOV I2,TMOV M2 WHERE M2.IDMOV=I2.IDMOV AND TITMMOVRELAC.IDMOVORIGEM=I2.IDMOV AND " & _
                           "TITMMOVRELAC.NSEQITMMOVORIGEM=I2.NSEQITMMOV AND TITMMOVRELAC.IDMOVDESTINO = M.IDMOV AND TITMMOVRELAC.NSEQITMMOVDESTINO = I.NSEQITMMOV AND M2.CODTMV='1.2.16' ),0) AS QTDEPENDENTE,(SELECT TVEN.CODVEN + ' - ' + TVEN.NOME FROM TVEN WHERE CODVEN=M.CODVEN2)AS NOMEQUEMEMPRESTOU,(SELECT TVEN.CODUSUARIO FROM TVEN WHERE CODVEN=M.CODVEN2)AS CODUSUARIORM,(SELECT CONVERT(VARCHAR(10),[RECMODIFIEDON],3)+' '+CONVERT(VARCHAR(5),[RECMODIFIEDON],108) FROM TITMMOV(NOLOCK) WHERE IDMOV=I.IDMOV AND NSEQITMMOV=I.NSEQITMMOV) AS DATAHORA FROM TITMMOV I (NOLOCK),TMOV M (NOLOCK),TPRODUTO P (NOLOCK),TLOC L (NOLOCK) WHERE I.CODCOLIGADA=M.CODCOLIGADA AND I.IDMOV=M.IDMOV AND  L.CODCOLIGADA=M.CODCOLIGADA AND L.CODFILIAL=M.CODFILIAL AND L.CODLOC=M.CODLOC AND " & _
                           "I.CODCOLIGADA=P.CODCOLPRD AND I.IDPRD=P.IDPRD AND M.CODTMV='2.2.15' AND M.STATUS<>'C' AND I.NSEQITMMOV NOT IN(SELECT NSEQITMMOV FROM TITMMOVBEM IB(NOLOCK) WHERE IB.IDMOV = I.IDMOV) AND (P.CODIGOPRD like '01.%' OR P.CODIGOPRD like '04.%' OR P.CODIGOPRD like '03.0001.1752' OR P.CODIGOPRD like '03.0001.1722' OR P.CODIGOPRD like '03.0001.1682' OR P.CODIGOPRD like '03.0001.3112' OR P.CODIGOPRD like '03.0001.3114') GROUP BY M.CODLOC,L.NOME,M.CODVEN1,M.NUMEROMOV,M.SERIE,M.SEQUENCIALESTOQUE,M.DATAEMISSAO,M.HORULTIMAALTERACAO,P.CODIGOPRD,P.NOMEFANTASIA,I.IDMOV,M.IDMOV,I.IDPRD,I.NSEQITMMOV,I.QUANTIDADE,M.CODVEN2,p.CAMPOLIVRE " & _
                           "hAVING I.QUANTIDADE - ISNULL((SELECT SUM(I2.QUANTIDADE) FROM TITMMOVRELAC(NOLOCK),TITMMOV I2,TMOV M2 wHERE M2.IDMOV=I2.IDMOV AND TITMMOVRELAC.IDMOVORIGEM=I2.IDMOV AND TITMMOVRELAC.NSEQITMMOVORIGEM=I2.NSEQITMMOV AND TITMMOVRELAC.IDMOVDESTINO=M.IDMOV AND TITMMOVRELAC.NSEQITMMOVDESTINO=I.NSEQITMMOV AND M2.CODTMV='1.2.16'),0) > 0 order by  codloc,I.IDMOV,P.CODIGOPRD asc"
    cnBancoSAP.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsBuscaDadosAntigos.Open sqlBuscaDadosAntigos, cnBancoSAP, adOpenKeyset, adLockReadOnly
    While Not rsBuscaDadosAntigos.EOF
        sqlCompletaDados = "select a.INATIVO,c.CODSITUACAO,a.NOME,c.CODFUNCAO,d.NOME,c.CODSECAO,e.DESCRICAO,a.CODUSUARIO from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b  on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA " & _
                           "left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO where a.CODCOLIGADA= 1 and a.codven = '" & rsBuscaDadosAntigos.Fields(2) & "' order by a.CODVEN"
        rsCompletaDados.Open sqlCompletaDados, cnBancoSAP, adOpenKeyset, adLockReadOnly
        If rsCompletaDados.RecordCount > 0 Then
            vNome = rsCompletaDados.Fields(2)
            If Not IsNull(rsCompletaDados.Fields(3)) Then vCodFuncao = rsCompletaDados.Fields(3)
            If Not IsNull(rsCompletaDados.Fields(4)) Then vNomeFuncao = rsCompletaDados.Fields(4)
            If Not IsNull(rsCompletaDados.Fields(5)) Then vCodSecao = rsCompletaDados.Fields(5)
            If Not IsNull(rsCompletaDados.Fields(6)) Then vNomeSecao = rsCompletaDados.Fields(6)
        End If
        rsCompletaDados.Close
        Set rsCompletaDados = Nothing
        
        
        sqlGravaEmprestimo = "select * from tbEmprestimo"
        rsGravaEmprestimo.Open sqlGravaEmprestimo, cnBanco, adOpenKeyset, adLockOptimistic
        rsGravaEmprestimo.AddNew
        rsGravaEmprestimo.Fields(0) = rsBuscaDadosAntigos.Fields(2)
        rsGravaEmprestimo.Fields(1) = vNome
        rsGravaEmprestimo.Fields(2) = vCodFuncao
        rsGravaEmprestimo.Fields(3) = vNomeFuncao
        rsGravaEmprestimo.Fields(4) = vCodSecao
        rsGravaEmprestimo.Fields(5) = vNomeSecao
        rsGravaEmprestimo.Fields(6) = Format(rsBuscaDadosAntigos.Fields(3), "yyyy/mm/dd")
        rsGravaEmprestimo.Fields(7) = rsBuscaDadosAntigos.Fields(16)
        rsGravaEmprestimo.Fields(8) = rsBuscaDadosAntigos.Fields(10)
        rsGravaEmprestimo.Fields(9) = rsBuscaDadosAntigos.Fields(11)
        rsGravaEmprestimo.Fields(10) = "E"
        rsGravaEmprestimo.Fields(11) = 1
        rsGravaEmprestimo.Fields(12) = Val(rsBuscaDadosAntigos.Fields(0))
        rsGravaEmprestimo.Fields(13) = rsBuscaDadosAntigos.Fields(22)
        rsGravaEmprestimo.Fields(14) = rsBuscaDadosAntigos.Fields(23)
        rsGravaEmprestimo.Update
        rsGravaEmprestimo.Close
        'Set rsGravaEmprestimo = Nothing

        vIDMovLocal = rsBuscaDadosAntigos(16)
        sqlGravaEmprestimoItens = "select * from tbEmprestimoItens"
        rsGravaEmprestimoItens.Open sqlGravaEmprestimoItens, cnBanco, adOpenKeyset, adLockOptimistic
        Do While rsBuscaDadosAntigos.Fields(16) = vIDMovLocal
            
            
            sqlCompletaDados = "select C.CODLOC,A.CODIGOPRD,B.CODUNDCONTROLE,A.NOMEFANTASIA,max(c.CUSTOMEDIO) CUSTOMEDIO,A.IDPRD from " & vBancoSAP & ".DBO.TPRODUTO AS A inner join " & vBancoSAP & ".DBO.TPRODUTODEF AS B on B.IDPRD=A.IDPRD inner join " & vBancoSAP & ".DBO.TPRDLOC AS C on C.IDPRD=A.IDPRD left join " & vBancoSAP & ".DBO.OFVENCPLANOMANUT AS D on D.IDOBJOF = SUBSTRING(A.CODIGOPRD,4,9) left join " & vBancoSAP & ".DBO.OFVENCPLANOMANUT AS E on E.IDOBJOF = D.IDOBJOF " & _
                               "left join " & vBancoSAP & ".DBO.OFPLANOMANUT AS F on F.IDPLANO = E.idplano and F.ATIVO = 1 where C.CODCOLIGADA=1 and A.INATIVO=0 and TIPO='P' and (A.CODIGOPRD like '01.%' or A.OBSERVACAO='FERRAMENTA' or A.CODIGOPRD like '04.%' OR A.CODIGOPRD like '03.0001.1752' OR A.CODIGOPRD like '03.0001.1722' OR A.CODIGOPRD like '03.0001.1682' OR A.CODIGOPRD like '03.0001.3112' OR A.CODIGOPRD like '03.0001.3114') and " & _
                               "A.NOMEFANTASIA like '%%' and C.CODLOC= '" & rsBuscaDadosAntigos.Fields(0) & "' and a.idprd = " & rsBuscaDadosAntigos.Fields(17) & " group by A.CODCOLPRD,A.NOMEFANTASIA,A.CODIGOPRD,C.SALDOFISICO2,C.SALDOFISICO6,B.CODUNDCONTROLE,A.IDPRD,B.PRECO1,C.CODLOC,D.DATAVENCIMENTO order by A.NOMEFANTASIA"
            rsCompletaDados.Open sqlCompletaDados, cnBancoSAP, adOpenKeyset, adLockReadOnly
            
            If rsCompletaDados.RecordCount > 0 Then
                rsGravaEmprestimoItens.AddNew
                rsGravaEmprestimoItens.Fields(0) = rsBuscaDadosAntigos.Fields(2)
                rsGravaEmprestimoItens.Fields(1) = rsBuscaDadosAntigos.Fields(0)
                rsGravaEmprestimoItens.Fields(4) = rsBuscaDadosAntigos.Fields(16)
                rsGravaEmprestimoItens.Fields(5) = rsBuscaDadosAntigos.Fields(17)
                rsGravaEmprestimoItens.Fields(6) = rsBuscaDadosAntigos.Fields(19)
                rsGravaEmprestimoItens.Fields(7) = rsBuscaDadosAntigos.Fields(20)
                rsGravaEmprestimoItens.Fields(2) = rsBuscaDadosAntigos.Fields(13)
                rsGravaEmprestimoItens.Fields(3) = rsBuscaDadosAntigos.Fields(14)
                rsGravaEmprestimoItens.Fields(8) = rsBuscaDadosAntigos.Fields(21)
                rsGravaEmprestimoItens.Fields(9) = Format(rsBuscaDadosAntigos.Fields(3), "yyyy/mm/dd")
                rsGravaEmprestimoItens.Fields(10) = Format(rsBuscaDadosAntigos.Fields(9), "hh:mm")
                rsGravaEmprestimoItens.Fields(11) = "E"
                rsGravaEmprestimoItens.Fields(12) = rsBuscaDadosAntigos.Fields(22)
                rsGravaEmprestimoItens.Fields(13) = rsBuscaDadosAntigos.Fields(18)
                rsGravaEmprestimoItens.Fields(14) = 1
                rsGravaEmprestimoItens.Fields(15) = rsCompletaDados.Fields(2)
                rsGravaEmprestimoItens.Fields(16) = rsBuscaDadosAntigos.Fields(21) * rsCompletaDados.Fields(4)
                rsGravaEmprestimoItens.Fields(17) = rsBuscaDadosAntigos.Fields(10)
                rsGravaEmprestimoItens.Fields(18) = rsBuscaDadosAntigos.Fields(11)
            End If
            rsCompletaDados.Close
            'Set rsCompletaDados = Nothing
            
            rsBuscaDadosAntigos.MoveNext
            If rsBuscaDadosAntigos.EOF Then Exit Do
        Loop
        rsGravaEmprestimoItens.Update
        rsGravaEmprestimoItens.Close
        Set rsGravaEmprestimoItens = Nothing
    Wend
    rsBuscaDadosAntigos.Close
    Set rsBuscaDadosAntigos = Nothing
    Exit Function
Err:
    If Err.Number = -2147217873 Then mobjMsg.Abrir "Dados já foram importados em outro momento", Ok, critico, "IMRM"
    Exit Function
End Function

Public Function enviaEmail()
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
'DICA: CRIA O DOCUMENTO NO WORD, COPIA, ABRE O OUTLOOK, CRIE UM NOVO EMAIL, COLE E SALVE COMO HTML
'VC IRÁ TRABALHAR COM O ARQUIVO HTML GERADO.
'SE O ARQUIVO FOR MUITO GRANDE, CRIE VÁRIAS VARIÁVEIS DO TIPO STRING PARA ARMAZENAR O ARQUIVO PICADO
'QUANDO FOR ENVIAR CONCATENE TODAS AS VARIAVEIS
On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    vValidaConf = 1
    
    vCorDecisao = "#CD2626"

    With Camp
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort (USAR CONFIGURAÇÕES DO OUTLOOK)
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPServerPort) = vPorta 'Informe a porta SMTP
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1 'vSMTPAutentic '1 '(MECANISMO DE AUTENTICAÇÃO)
        .Item(cdoSMTPUseSSL).Value = vSSL '(INFORMA SE REQUER CONFIGURAÇÃO SEGURA)
        .Item(cdoSendUserName) = "guilhermemfa@gmail.com" ' informe o usuario de autenticação
        .Item(cdoSendPassword) = "soeu2008"  'Informe a Senha aqui
        .Update
    
'-------- CONFIGURAÇÃO ORIGINAL VIGA
'        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
'        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
'        .Item(cdoSMTPServerPort) = vPorta 'Informe a porta SMTP
'        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
'        .Item(cdoSMTPAuthenticate) = 1
'        .Item(cdoSMTPUseSSL).Value = False
'        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
'        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
'        .Update
    
    End With

    With Msg
        Set .Configuration = Cof
        .To = "guilhermemfa@gmail.com" 'destinatarios separados por ;
        .From = "guilhermemfa@gmail.com"  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "Assunto: teste de envio de email"
        
        .HTMLBody = "Teste de envio de email"
       .Send
    End With
    Exit Function
errMail:
    Msgbox "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SAF"
    vValidaConf = 0
    Exit Function
End Function


