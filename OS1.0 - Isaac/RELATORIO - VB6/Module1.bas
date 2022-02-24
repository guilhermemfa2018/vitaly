Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long 'Biblioteca para manipulação do Regedit
Public cnBancoDBF As ADODB.Connection
Public sDatabaseName As String
'Public sCompanyName As String
Public vOS As String

Public Function conexaoDBF()
On Error GoTo Err
    conexaoDBF = True
    Set cnBancoDBF = New ADODB.Connection
    cnBancoDBF.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabaseName & ";Extended Properties=""DBASE IV;"";"
    cnBancoDBF.Open
    frmGeraRelatório.Label1.ForeColor = &H8000&
    frmGeraRelatório.Label1.Caption = "Conectado com sucesso!!!"
    Exit Function
Err:
    conexaoDBF = False
    frmGeraRelatório.Label1.ForeColor = &HC0&
    frmGeraRelatório.Label1.Caption = "Falha de conexao"
    MsgBox Err.Description
    Exit Function
End Function

Public Function gravaCaminhoNoRegedit()
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\" & "sDatabaseName", sDatabaseName 'Chave com o nome do Banco
    'Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\" & "sCompanyName", sCompanyName 'Chave com o nome do Banco
End Function


Public Function carregaCaminhoDoRegedit()
On Error GoTo Err1
    Set Reg = CreateObject("wscript.shell")
    If Reg.regread("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\sDatabaseName") <> "" Then
        frmGeraRelatório.txtIntegra(10).Text = Reg.regread("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\sDatabaseName")
        sDatabaseName = frmGeraRelatório.txtIntegra(10).Text
        'frmGeraRelatório.txtIntegra(11).Text = Reg.regread("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\sCompanyName")
        'sCompanyName = frmGeraRelatório.txtIntegra(11).Text
    End If
    Exit Function
Err1:
    'Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\" & "sDatabaseName", sDatabaseName 'Chave com o nome do Banco
    'Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\DB_ICS\" & "sCompanyName", sCompanyName 'Chave com o nome do Banco
    'MsgBox Err.Description
End Function


Function CarRemove(Texto As String, Caracteres As String) As String
    Dim Ic As Integer, It As Integer, Inicio As Integer, Pos As Integer, Caracter As String
    
    For Ic = 1 To Len(Caracteres)
        Caracter = Mid(Caracteres, Ic, 1)
        Pos = 1
        Inicio = 1
        If InStr(Inicio, Texto, Caracter) > 0 Then
            For It = 1 To Len(Texto)
                Pos = InStr(Inicio, Texto, Caracter)
                CarRemove = Mid(Texto, 1, Pos - 1) & Mid(Texto, Pos + 1)
                Inicio = Pos
                'List1.AddItem CarRemove & [~]|[~] & Pos & Caracter
            Next It
            Texto = CarRemove
        Ic = Ic - 1
        Else
            CarRemove = Texto
        End If
    Next Ic
End Function
