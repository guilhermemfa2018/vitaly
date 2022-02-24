Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Biblioteca para manipulação do Regedit
Public colheDados(17) As String 'Guarda dados de importação de colaboradores de arquivo TXT
Public cnBanco As ADODB.Connection
Public cnBancoFlexJunior As ADODB.Connection
Public cnBancoZeus As ADODB.Connection
Public vFornec As String
Public vProduto As String
Public vFCE As String
Public vDataFilter1 As String
Public vDataFilter2 As String
Public vMovimento As String
Public vMovs As String
Public vCustos As String
Public vFCECC As String
Public vColigada As Integer
Public vComando As String

'VARIAVEIS PARA ARMAZENAR DADOS DE CONEXAO DO BANCO FDB (FLEXJR)
Public vUsuarioFlexJr As String
Public vPassFlexJr As String
Public vPathDBFlexJr As String

'VARIAVEIS PARA ARMAZENAR DADOS DE CONEXAO COM O RELÓGIO DE PONTO
Public vIDRelogio As String
Public vIPRelogio As String
Public vCPFResponsavel As String
Public vPassRelogio As String
Public vCaminhoDadosCapturadoRelogio As String

'ABAIXO CONEXÃO COM O BANCO DE DADOS
Public Function Conectar(vServer As String, vDB As String, vUser As String, vPass As String)
On Error GoTo Err1
    Conectar = True
    
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & vPass & ";Persist Security Info=True;Connect Timeout=0;User ID=" & vUser & ";Initial Catalog=" & vDB & ";Data Source=" & vServer

    Exit Function
Err1:
    Conectar = False
    Exit Function
End Function

Public Function conexaoFB()
On Error GoTo Err
    frmBatidaDiaria.Label1.Caption = "Conectando com FlexJr"
    conexaoFB = True
    Set cnBancoFlexJunior = New ADODB.Connection
'    cnBancoFlexJunior.ConnectionString = "DRIVER=Firebird/InterBase(r) driver;UID=" & vUsuarioFlexJr & ";PWD=" & vPassFlexJr & ";DBNAME=" & vPathDBFlexJr
    cnBancoFlexJunior.ConnectionString = "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=C:\ZEUS\PONTO\FLEXJR.FDB"
    cnBancoFlexJunior.Open
    Exit Function
Err:
    conexaoFB = False
    frmBatidaDiaria.Label1.Caption = Err.Description
    MsgBox Err.Description
    Exit Function
End Function

Public Function carregaDadosConexaoRelogioPonto()
    Dim rsCarregaDadosConexaoRelogioPonto As New ADODB.Recordset
    Dim SqlCarregaDadosConexaoRelogioPonto  As String
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "SELECT " & vbCrLf
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "   IDRELOGIO, " & vbCrLf
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "   IPRELOGIO, " & vbCrLf
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "   CPFRESPONSAVEL, " & vbCrLf
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "   PASSWORD, " & vbCrLf
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "   CAMINHO " & vbCrLf
    SqlCarregaDadosConexaoRelogioPonto = SqlCarregaDadosConexaoRelogioPonto & "FROM TBCONFRELOGIO"
    rsCarregaDadosConexaoRelogioPonto.Open SqlCarregaDadosConexaoRelogioPonto, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsCarregaDadosConexaoRelogioPonto.RecordCount > 0 Then
        vIDRelogio = rsCarregaDadosConexaoRelogioPonto.Fields(0)
        vIPRelogio = rsCarregaDadosConexaoRelogioPonto.Fields(1)
        vCPFResponsavel = rsCarregaDadosConexaoRelogioPonto.Fields(2)
        vPassRelogio = rsCarregaDadosConexaoRelogioPonto.Fields(3)
        vCaminhoDadosCapturadoRelogio = rsCarregaDadosConexaoRelogioPonto.Fields(4)
    End If
    rsCarregaDadosConexaoRelogioPonto.Close
End Function


Public Function carregaDadosConexaoFlexJr()
    Dim rsCarregaDadosConexaoFlexJr As New ADODB.Recordset
    Dim SqlCarregaDadosConexaoFlexJr  As String
    
    SqlCarregaDadosConexaoFlexJr = SqlCarregaDadosConexaoFlexJr & "SELECT " & vbCrLf
    SqlCarregaDadosConexaoFlexJr = SqlCarregaDadosConexaoFlexJr & "   USUARIO, " & vbCrLf
    SqlCarregaDadosConexaoFlexJr = SqlCarregaDadosConexaoFlexJr & "   PASSWORD, " & vbCrLf
    SqlCarregaDadosConexaoFlexJr = SqlCarregaDadosConexaoFlexJr & "   CAMINHO " & vbCrLf
    SqlCarregaDadosConexaoFlexJr = SqlCarregaDadosConexaoFlexJr & "FROM TBCONFFLEXJR"
    rsCarregaDadosConexaoFlexJr.Open SqlCarregaDadosConexaoFlexJr, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsCarregaDadosConexaoFlexJr.RecordCount > 0 Then
        vUsuarioFlexJr = rsCarregaDadosConexaoFlexJr.Fields(0)
        vPassFlexJr = rsCarregaDadosConexaoFlexJr.Fields(1)
        vPathDBFlexJr = rsCarregaDadosConexaoFlexJr.Fields(2)
    End If
    rsCarregaDadosConexaoFlexJr.Close
End Function

Public Function RemoveMask(campo)
    Dim Variavel As String
    Dim Varival As String
    Variavel = Replace(campo, "/", "")
    RemoveMask = Variavel
End Function

Public Function formatData(vData As String)
    Dim vDataFormatada As String
    vDataFormatada = Mid$(vData, 1, 2) & "-" & Mid$(vData, 3, 2) & "-" & Mid$(vData, 5, 4)
    vDataFormatada = Format(vDataFormatada, "yyyy-mm-dd")
    formatData = vDataFormatada
End Function

Public Function formatHora(vHora As String)
    Dim vHoraFormatada As String
    vHoraFormatada = Mid$(vHora, 1, 2) & ":" & Mid$(vHora, 3, 2)
    formatHora = vHoraFormatada
End Function
