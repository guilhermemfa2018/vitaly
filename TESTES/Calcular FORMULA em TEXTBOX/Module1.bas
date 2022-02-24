Attribute VB_Name = "Module1"
Public cnBanco As ADODB.Connection


'ABAIXO CONEXÃO COM O BANCO DE DADOS
Public Function Conectar()
On Error GoTo Err1
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & Form1.Text4.Text & ";Persist Security Info=True;Connect Timeout=0;User ID=" & Form1.Text5.Text & ";Initial Catalog=" & Form1.Text6.Text & ";Data Source=" & Form1.Text7.Text
    Exit Function
Err1:
    Debug.Print Err.Number + " - " + Err.Description
    Exit Function
End Function
