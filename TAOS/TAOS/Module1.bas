Attribute VB_Name = "Module1"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public cnBanco As ADODB.Connection
Public rsLocal As New ADODB.Recordset
Public Sqlp As String
Public procnom As String, procnom1 As String
Public vMSGGlobal As String
Public vNomeGlobal As String
Public vVerificaPermissao As Integer
Public vSMTP As String
Public vUsuEmail As String
Public vSenhaEmail As String
Public sEmailSI As String 'String que guarda endereços de e-mail que receberão notificações SI - Solicitação de Inspeção
Public sEmailSRM As String 'String que guarda endereços de e-mail que receberão notificações SRM - Solicitação de Retirada de Material



'Public Pesquisa As String
'Public campo As Integer
'Public Campo1 As Integer
'Public campo2 As Integer
'Public campo3 As Integer
'Public Campo4 As Integer
'Public vQualquerDado(50, 30) As String

'ABAIXO CONEXÃO COM O BANCO DE DADOS
Public Function Conectar()
'On Error GoTo Err1
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & Form1.Text4.Text & ";Persist Security Info=True;Connect Timeout=0;User ID=" & Form1.Text5.Text & ";Initial Catalog=" & Form1.Text6.Text & ";Data Source=" & Form1.Text7.Text
    'Form1.Label20.Visible = False
    'Form1.Visible = False
    msgLabel "Conexão Reestabelecida", 1
    Exit Function
Err1:
    'Form1.Label20.Visible = True
    'Form1.WindowState = 0 ' normal
    Exit Function
End Function

'A Função abaixo LIMPA DADOS de qualquer ListView
Public Function LimpaLV(LV As ListView)
    LV.ListItems.Clear
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
                ItemLst.SubItems(X) = "" & rsCompoe.Fields(X)
            End If
        Next
        rsCompoe.MoveNext
    Wend
    LV.Sorted = True
    LV.SortKey = 0
    LV.SortOrder = lvwAscending
    rsCompoe.Close
    Set rsCompoe = Nothing
End Function

Function AlwaysOnTop(FrmID As Form, ByVal OnTop As Boolean) As Boolean
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop = True Then
        AlwaysOnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        AlwaysOnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Function

Public Function msgLabel(vMensagem As String, vTipo As Integer)
    If vTipo = 1 Then Form1.Label3.ForeColor = &H8000& 'Verde
    If vTipo = 2 Then Form1.Label3.ForeColor = &HC0& 'Vermelho
    Form1.Label3.Caption = vMensagem
    If vTipo = 1 Then
        vSetFocus = 1
        Form1.Nok.Visible = False
        If vMensagem = "" Then
            Form1.Ok.Visible = False
        Else
            Form1.Ok.Visible = True
        End If
        Form1.Label3.FontSize = 14
        Form1.Label3.FontBold = False
    End If
    If vTipo = 2 Then
        vSetFocus = 2
        Form1.Nok.Visible = True
        Form1.Ok.Visible = False
        Form1.Label3.FontSize = 18
        Form1.Label3.FontBold = True
    End If
End Function

Public Function carregaDadosEmail()
    Dim rsConfEmail As New ADODB.Recordset
    Dim sqlConfEmail As String

    sqlConfEmail = "Select * from tbConfEmail where codcoligada = 5 " ''" & vCodcoligada & "'"
    rsConfEmail.Open sqlConfEmail, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsConfEmail.EOF Then
        vSMTP = rsConfEmail.Fields(0)
        vUsuEmail = rsConfEmail.Fields(1)
        vSenhaEmail = rsConfEmail.Fields(2)
    End If
    rsConfEmail.Close
End Function
