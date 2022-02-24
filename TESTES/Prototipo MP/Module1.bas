Attribute VB_Name = "Module1"
Public cnBanco As ADODB.Connection
Public rsLocal As New ADODB.Recordset
Public Sqlp As String
Public procnom As String, procnom1 As String
Public Pesquisa As String
Public campo As Integer
Public Campo1 As Integer
Public campo2 As Integer
Public campo3 As Integer
Public Campo4 As Integer
Public vQualquerDado(50, 30) As String

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

'A Função abaixo EXCLUI linhas de quaisquer ListView
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

'A Função abaixo gera código para qualquer Listview
Public Function GeraCodigoLV(LV As ListView)
    If LV.ListItems.Count > 0 Then
        Dim X As Integer
        X = 1
        LV.Sorted = True
        LV.SortOrder = lvwDescending
        LV.ListItems.Item(X).Selected = True
        GeraCodigoLV = LV.ListItems.Item(X) + 1
        LV.SortOrder = lvwAscending
        Exit Function
    Else
        GeraCodigoLV = 1
    End If
End Function

'A Função abaixo gera código para qualquer Tabela
Public Function GeraCodigoTB(vTabela As String, vCampo As String)
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    SqlGera = "Select top 1 * from " & vTabela & " order by " & vCampo & " Desc"
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
    If vTabela = "tbGrupoClass" Then
        Sqlp = "Select " & vPesq1 & "," & vPesq2 & " from " & vTabela & " where idprd ='" & frmFormulaCC.txtformula(0) & "' order by " & vCampo & ""
    Else
        Sqlp = "Select " & vPesq1 & "," & vPesq2 & " from " & vTabela & "  order by " & vCampo & ""
    End If
    procnom = vCampo
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa"
    Pesquisa = vForm.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
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

'A Função abaixo é referente a ALTERAÇÃO de dados de qualquer ListView com até 10 colunas
Public Function AlteraLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox)
    Dim Y As Integer, X As Integer, Z As Integer
    Dim vRaptor(10) As String
    For X = LBound(vRaptor) To UBound(vRaptor)
        vRaptor(X) = ""
    Next X
    Y = LV.ListItems.Count
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    
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
    If vRaptor(10) <> "" Then vCP10.Text = vRaptor(10)
End Function

'A Função abaixo é referente a INCLUSÃO de dados de qualquer ListView com até 10 colunas
Public Function IncluirLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox)
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, Z As Integer
    Dim vRaptor(10) As TextBox
    'For X = LBound(vRaptor) To UBound(vRaptor)
    '    vRaptor(X) = ""
    'Next X
    
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
    If ValidaCampos(LV, vCP01, vCP02, vCP03, vCP04, vCP05, vCP06, vCP07, vCP08, vCP09, vCP10) = False Then Exit Function
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
                'LimpaControles vCP01, vCP02, vCP03, vCP04, vCP05, vCP06, vCP07, vCP08, vCP09, vCP10
                Exit Function
            End If
        Next
        Set ItemLst = LV.ListItems.Add(, , vRaptor(1))
        Y = LV.ListItems.Count
    Else
        Set ItemLst = LV.ListItems.Add(, , vRaptor(1))
        Y = LV.ListItems.Count
    End If
    For Z = 2 To LV.ColumnHeaders.Count
        If vRaptor(Z) <> "" Then ItemLst.SubItems(Z - 1) = vRaptor(Z)
    Next
    'LimpaControles vCP01, vCP02, vCP03, vCP04, vCP05, vCP06, vCP07, vCP08, vCP09, vCP10
    If vRaptor(2).Visible = True Then vRaptor(2).SetFocus
End Function

'A Função abaixo é referente a VALIDAÇÃO de dados de qualquer ListView com até 10 colunas
Public Function ValidaCampos(LV As ListView, vTxt1 As TextBox, vTxt2 As TextBox, vTxt3 As TextBox, vTxt4 As TextBox, vTxt5 As TextBox, vTxt6 As TextBox, vTxt7 As TextBox, vTxt8 As TextBox, vTxt9 As TextBox, vTxt10 As TextBox)
    On Error Resume Next
    ValidaCampos = False
    Dim X As Integer
    Dim vMatrix(10) As TextBox
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
    For X = 1 To LV.ColumnHeaders.Count
        If vMatrix(X) = "" Then
            MsgBox "Favor informar o campo: " & vMatrix(X).Tag, vbInformation, "Atenção"
            vMatrix(X).SetFocus
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
Public Function CarregaTxt(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vCampo2 As String, vTipoCampo2 As String, vVar1 As TextBox, vVar2 As TextBox, vPosicao1 As Integer, vPosicao2 As Integer, vRetorno1 As TextBox, vTipoRetorno1 As String, vRetorno2 As TextBox)
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
    
    Dim X As Integer
    Dim rsCarregaTxT As New ADODB.Recordset
    Dim sqlCarregaTxt As String
    
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
    rsCarregaTxT.Open sqlCarregaTxt, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCarregaTxT.EOF Then rsCarregaTxT.MoveFirst
    If rsCarregaTxT.EOF Then
        MsgBox "Dado não cadastrado", vbCritical, "PrototipoX"
    Else
        If vTipoRetorno1 = "S" Then
            vRetorno1.Text = rsCarregaTxT.Fields(vPosicao1)
        Else
            vRetorno1.Text = Format(rsCarregaTxT.Fields(vPosicao1), "00")
        End If
        vRetorno2.Text = rsCarregaTxT.Fields(vPosicao2)
    End If
    rsCarregaTxT.Close
    Set rsCarregaTxT = Nothing
End Function

'A Função abaixo grava dados de até 50 variáveis de um formulário em uma determinada tabela
Public Function GravaDados(vTabela As String, vCampo1 As String, vTipoCampo1 As String, vVar1 As TextBox, vQtdCampos As Integer)
    'vTabela     = Nome da tabela a qual será realizada a pesquisa da Query
    'vCampo1     = Nome do campo da 1ª condição de pesquisa da Query
    'vTipoCampo1 = Tipo do 1º campo de pesquisa da Query
    'vVar1       = Nome do 1º TextBox que contem o valor que será pesquisado no 1ª campo de pesquisa da Query
    'vQtdCampos  =  Quantidade de variáveis que serão gravados na tabela
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer, X As Integer
    cnBanco.BeginTrans
   
    If vTipoCampo1 = "I" Then
        SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & Val(vVar1) & "'"
    Else
        SqlSalvar = "select * from " & vTabela & " where " & vCampo1 & " = '" & vVar1 & "'"
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
    cnBanco.CommitTrans
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

Public Function ordenaLVArray(LV As ListView, vPos0 As String, vPos1 As String, vPos2 As String, vPos3 As String, vPos4 As String, vPos5 As String, vPos6 As String, vPos7 As String, vPos8 As String, vPos9 As String, vPos10 As String)
    Dim X As Integer, Y As Integer, Z As Integer
    Dim vMatrix(11) As String
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
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        For Z = 0 To LV.ColumnHeaders.Count
            If IsNumeric(vMatrix(Z + 1)) Then
                If vMatrix(Z + 1) = "0" Then
                    vQualquerDado(X, Z + 1) = LV.ListItems.Item(X)
                Else
                    vQualquerDado(X, Z + 1) = LV.SelectedItem.ListSubItems.Item(Val(vMatrix(Z + 1)))
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
'On Error GoTo TrataErro
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer, X As Integer
    
    cnBanco.BeginTrans
    
    If vCampo1 = "" Then
        sqlDeletar = "Delete from " & vTabela
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
        While vQualquerDado(X, Y) <> ""
            rsSalvar.Fields(Y - 1) = vQualquerDado(X, Y)
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
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbCritical, "Atenção"
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

Public Function mudaCorText(txt As TextBox)
    txt.BackColor = &HC0E0FF
End Function

Public Function voltaCorText(txt As TextBox)
    txt.BackColor = &HFFFFFF
End Function
