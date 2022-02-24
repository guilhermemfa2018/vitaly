Attribute VB_Name = "Module1"
Public cnBanco As ADODB.Connection
Public cnBanco2 As ADODB.Connection
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

'ABAIXO CONEXÃO COM O BANCO DE DADOS
Public Function Conectar()
'On Error GoTo Err1
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & Form1.Text4.Text & ";Persist Security Info=True;Connect Timeout=0;User ID=" & Form1.Text5.Text & ";Initial Catalog=" & Form1.Text6.Text & ";Data Source=" & Form1.Text7.Text
    'Form1.Label20.Visible = False
    'Form1.Visible = False
    Exit Function
Err1:
    'Form1.Label20.Visible = True
    'Form1.WindowState = 0 ' normal
    Exit Function
End Function


Public Function Conectar2()
'On Error GoTo Err1
    Set cnBanco2 = New ADODB.Connection
    cnBanco2.Open "Provider=SQLOLEDB.1;Password=" & Form1.Text4.Text & ";Persist Security Info=True;User ID=" & Form1.Text5.Text & ";Data Source=" & Form1.Text7.Text
    'Form1.Label20.Visible = False
    'Form1.Visible = False
    Exit Function
Err1:
    'Form1.Label20.Visible = True
    'Form1.WindowState = 0 ' normal
    Exit Function
End Function

