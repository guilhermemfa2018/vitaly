Attribute VB_Name = "Module1"
Public cnBanco As ADODB.Connection
Public cnBancoSAP As ADODB.Connection
Public oConn As ADODB.Connection
Public rsLocal As New ADODB.Recordset
Public Sqlp As String
Public vDataFilter1 As String
Public vDataFilter2 As String
Public vBancoTotvs As String  'Armazena nome do banco totvs

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
                        If vRaptor(Z) <> "" Then vRaptor(Z) = Format(LV.ListItems.Item(X), "000")
                    Else
                        If vRaptor(Z) <> "" Then LV.SelectedItem.ListSubItems.Item(Z - 1) = vRaptor(Z)
                    End If
                Next
                Y = LV.ListItems.Count
                IncluirLV = True
                Exit Function
            End If
        Next

        Set ItemLst = LV.ListItems.Add(, , Format(vRaptor(1), "00000"))
        Y = LV.ListItems.Count
    Else
        Set ItemLst = LV.ListItems.Add(, , Format(vRaptor(1), "000"))
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

'Essa rotina serve para verificar se o item/c.custo que esta sendo inserido no ListView1
'Ja está em uma outra OS
Private Function separaDesLv(vTxtForm As String)
    separaDesLv = True
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    Dim vCodLM As String, vCodSeq As String
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If Len(RECEBE) = 5 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 3)
            Else
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 4)
            End If
'            vCodLM = Mid$(RECEBE, 1, 2)
'            vCodSeq = Mid$(RECEBE, 3, 3)
            SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
            "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoTotvs & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
            "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoTotvs & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
            "inner join tbProjetos as g on g.codprojeto = d.codprojeto where a.fce = '" & Val(chamaForm.txtformula(12)) & "' and a.projeto = '" & chamaForm.txtformula(13).Text & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "' and a.idcc = '" & chamaForm.txtformula(0) & "' and a.idoperacao ='" & chamaForm.Combo1 & "'"
            rsTransf.Open SqlTransf, cnBanco, adOpenKeyset, adLockReadOnly
            If rsTransf.RecordCount > 0 Then
                mobjMsg.Abrir "Desenho: " & rsTransf.Fields(8) & vbCrLf & _
                              "Posição: " & rsTransf.Fields(11) & vbCrLf & _
                              "Item:" & rsTransf.Fields(12) & vbCrLf & _
                              "C.Custo:" & rsTransf.Fields(6) & vbCrLf & _
                              "Registrado na OS:" & Format(rsTransf.Fields(0), "000000000") & " - Programação: " & Format(rsTransf.Fields(7), "000000"), Ok, critico, "Atenção"
                separaDesLv = False
                rsTransf.Close
                Set rsTransf = Nothing
                Exit Function
            End If
            rsTransf.Close
            Set rsTransf = Nothing
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If Len(RECEBE) = 5 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 3)
        Else
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 4)
        End If
'        vCodLM = Mid$(RECEBE, 1, 2)
'        vCodSeq = Mid$(RECEBE, 3, 3)
        SqlTransf = "select a.idos,a.revisao,a.fce,a.projeto,a.codlm,a.codseq,a.idcc,a.idprogramacao,d.desenho,d.revisao,c.NOMEFANTASIA,e.posicao,e.item from tbositens as a " & _
        "inner join tbItemLM as b on a.fce = b.fce and a.codlm = b.codlm and a.codseq = b.codseq inner join " & vBancoTotvs & ".dbo.tprd as c on b.codmat = c.IDPRD " & _
        "inner join tbDesenhos as d on b.codigodes = d.iddesenho inner join tbPosicoes as e on b.codigopos = e.codigopos left join " & vBancoTotvs & ".dbo.TTB2 as f on c.CODTB2FAT = f.CODTB2FAT " & _
        "inner join tbProjetos as g on g.codprojeto = d.codprojeto where a.fce = '" & Val(chamaForm.txtformula(12)) & "' and a.projeto = '" & chamaForm.txtformula(13).Text & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "' and a.idcc = '" & chamaForm.txtformula(0) & "' and a.idoperacao ='" & chamaForm.Combo1 & "'"
        rsTransf.Open SqlTransf, cnBanco, adOpenKeyset, adLockReadOnly
        If rsTransf.RecordCount > 0 Then
            mobjMsg.Abrir "Desenho: " & rsTransf.Fields(8) & vbCrLf & _
                          "Posição: " & rsTransf.Fields(11) & vbCrLf & _
                          "Item:" & rsTransf.Fields(12) & vbCrLf & _
                          "C.Custo:" & rsTransf.Fields(6) & vbCrLf & _
                          "Registrado na OS:" & Format(rsTransf.Fields(0), "000000000") & " - Programação: " & Format(rsTransf.Fields(7), "000000"), Ok, critico, "Atenção"
            separaDesLv = False
            rsTransf.Close
            Set rsTransf = Nothing
            Exit Function
        End If
        rsTransf.Close
        Set rsTransf = Nothing
    End If
End Function

Public Sub Conexao()
'On Error GoTo Err1
    'If sServerName = "" Then GoTo Err1
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & frmimportarnfe.Text8.Text & ";Persist Security Info=True;User ID=" & frmimportarnfe.Text5.Text & ";Initial Catalog=" & frmimportarnfe.Text4.Text & ";Data Source=" & frmimportarnfe.Text3.Text
    frmimportarnfe.Label5.ForeColor = &H8000&
    frmimportarnfe.Label5.Caption = "Conexão com Banco Totvs RM realizada com sucesso"
    Exit Sub
Err1:
    frmimportarnfe.Label5.ForeColor = &HFF&
    frmimportarnfe.Label5.Caption = "Falha na Conexão.Erro ao tentar acessar DB - Entre com as novas configurações do servidor "
    Exit Sub
End Sub

Public Function CriarTabelasADO() As Boolean
On Error GoTo Err1
    
    'CRIA TABELAS SAF
    
    oConn.Execute "CREATE TABLE " & frmimportarnfe.Text4.Text & ".dbo.tbNFE(" & _
    "id INT NOT NULL identity (1,1)," & _
    "nfe VARCHAR(35) NOT NULL," & _
    "serie VARCHAR(5) NOT NULL," & _
    "cnpj VARCHAR(20) NULL," & _
    "fornecedor VARCHAR(200) NULL," & _
    "dtemissao DATETIME NULL," & _
    "dtentrada DATETIME NULL," & _
    "valornf FLOAT NULL," & _
    "chavenf VARCHAR(200) NULL," & _
    "dtcadastro DATETIME NULL," & _
    "codcoligada INT NULL," & _
    "PRIMARY KEY (id))"
    
    oConn.Execute "CREATE TABLE " & frmimportarnfe.Text4.Text & ".dbo.tbLogoColigada(" & _
    "id INT NOT NULL identity (1,1)," & _
    "codcoligada INT NOT NULL," & _
    "caminhoLogo VARCHAR(300) Not NULL," & _
    "PRIMARY KEY (id))"
    
    frmimportarnfe.Label5.ForeColor = &H8000&
    frmimportarnfe.Label5.Caption = "Tabela criada com sucesso"
    Exit Function
Err1:
    'Msgbox "(ADO) Erro ao criar Tabela de dados: " & vbCrLf & Err.Number & " - Tabela já Existe - " & Err.Description, 16, "Mensagem de erro"
    Resume Next
    'Exit Function
End Function

Public Function CriarBancoDeDadosADO() As Boolean
On Error GoTo Err1
    
    Set oConn = New ADODB.Connection
    
    ' Cria Banco
    oConn.Open "Provider=SQLOLEDB;Data Source=" & frmimportarnfe.Text3.Text & ";User ID=" & frmimportarnfe.Text5.Text & ";Password=" & frmimportarnfe.Text8.Text & ";"
    oConn.Execute "CREATE DATABASE " & frmimportarnfe.Text4.Text
    
    
    oConn.Close
    Set oConn = Nothing
    
    'MsgBox "Banco criado com sucesso", vbInformation, "SAF"
    Exit Function
Err1:
    frmimportarnfe.Label5.ForeColor = &HFF&
    frmimportarnfe.Label5.Caption = "(ADO) Erro ao criar banco de dados: " & vbCrLf & Err.Number & " - DB já Existe - " & Err.Description
    Exit Function
End Function

Public Sub CompoeCombo1(Combo As ComboBox, Tabela, campo, Campo1)
    Dim sql As String
    Dim rsTabela As New ADODB.Recordset
    Dim X As Integer
'    sql = "Select * from " & Tabela & " where codcoligada = '" & vCodcoligada & "' Order By " & campo
    sql = "Select * from " & Tabela & " where codcoligada > 0 "
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsTabela.EOF() Then
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

Public Function RemoveMask(campo)
    Dim Variavel As String
    Dim Varival As String
    Variavel = Replace(campo, ".", "")
    Variavel = Replace(Variavel, "-", "")
    Variavel = Replace(Variavel, "/", "")
    RemoveMask = Variavel
End Function


'A Função abaixo é referente a ALTERAÇÃO de dados de qualquer ListView com até 15 colunas
Public Function AlteraLV(LV As ListView, vCP01 As TextBox, vCP02 As TextBox, vCP03 As TextBox, vCP04 As TextBox, vCP05 As TextBox, vCP06 As TextBox, vCP07 As TextBox, vCP08 As TextBox, vCP09 As TextBox, vCP10 As TextBox, vCP11 As TextBox, vCP12 As TextBox, vCP13 As TextBox, vCP14 As TextBox, vCP15 As TextBox)
    Dim Y As Integer, X As Integer, Z As Integer
    Dim vRaptor(15) As String
    For X = LBound(vRaptor) To UBound(vRaptor)
        vRaptor(X) = ""
    Next X
    Y = LV.ListItems.Count
    If Y = 0 Then Exit Function
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    
    'SOMENTE PARA ZEUS--
    If apontaLV = 9 Then
        '1º VERIFICA SE A DATA PREVISTA ESTA VAZIA
        If LV.SelectedItem.ListSubItems(5).Text <> "" And LV.SelectedItem.ListSubItems(5).Text <> "-" Then
            '2º VERIFICA SE A SEMANA ATUAL É MAIOR OU IGUAL A SEMANA PROGRAMADA
            If DatePart("ww", (Date)) >= DatePart("ww", CDate(LV.SelectedItem.ListSubItems(5).Text)) Then
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
    '-------------------
    
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
    If vRaptor(10) <> "" Then vCP10.Text = vRaptor(10) 'Else vCP10.Text = ""
    If vRaptor(11) <> "" Then vCP11.Text = vRaptor(11) 'Else vCP11.Text = ""
    If vRaptor(12) <> "" Then vCP12.Text = vRaptor(12) 'Else vCP12.Text = ""
    If vRaptor(13) <> "" Then vCP13.Text = vRaptor(13) 'Else vCP13.Text = ""
    If vRaptor(14) <> "" Then vCP14.Text = vRaptor(14) 'Else vCP14.Text = ""
    If vRaptor(15) <> "" Then vCP15.Text = vRaptor(15) 'Else vCP15.Text = ""
End Function
