Attribute VB_Name = "CompoeLVs"
Public vDataFilter1 As String
Public vDataFilter2 As String
Public apontaLV As Integer
Public indiceVarGlobal As Integer 'quantas colunas vai ter a variavel global
Public checaFiltro As Boolean
Public vADP(10, 1) As String
Public diasTrabalhados As Integer
Public vQdtFrom As Integer 'Especifica a quantidade de FROM na query do filtro
Public tipoADP As String
Public vTimer As Boolean
Public vListViewTipoMaterial As Listview, vListViewParadas As Listview, vListviewClientes As Listview, vListviewTransportadoras As Listview, vListviewFormulaPRD As Listview, vListviewComercial As Listview, vListviewFormulaCC As Listview, vListviewDesenhos As Listview, vListviewFaturamentoFCE As Listview, vListviewFCE As Listview, vListviewLM As Listview, vListviewMP As Listview, vListviewControleDesenhos As Listview, vListviewRNCF As Listview, vListviewRelInsp As Listview, vListviewImpInspecao As Listview, vListviewRelExpedicao As Listview, vListviewImpExpedicao As Listview, vListviewGrupos As Listview, vListviewUsuarios As Listview, vListviewPermissoes As Listview, vListviewTerceiros As Listview

Public Function montaLV1(QualLV As Integer)
    On Error GoTo TrataErro
    
    If vAvisos = "" Then
        Msgbox "Local de Estoque não ativo. Acesse: Configurações|Sistema|Parametrizações|Gerais e informe", vbCritical, "Zeus"
        Exit Function
    ElseIf vBancoTotvs = "" Then
        Msgbox "Parâmetros de integração não informados. Acesse: Configurações|Sistema|Parametrizações|Integração e informe", vbCritical, "Zeus"
        Exit Function
    ElseIf vCodcoligada = 0 Then
        Msgbox "Coligada não cadastrada. Acesse: Configurações|Sistema|Coligadas e informe", vbCritical, "Zeus"
        Exit Function
    End If
    
    montaLV1 = True
    vTimer = False
    If Pesquisa <> "filtro" Then
    End If
    Permissao
    frmPesqGeralTeste2.SSTab1.Caption = Formulario

    QtdColReal = 0

'-- TIPO DE MATERIAIS
    contruirBotoesPorModulo QualLV
    If QualLV = 0 Then
        Set vListViewTipoMaterial = vListViewPrincipal
        Set chamaForm = New frmTipoMat
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 2
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then frmFiltro.Show 1
            carregaTABS "tbTipoMat", "", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then SqlLV = "SELECT A.CODIGO,A.DESCRICAO,A.ATIVO FROM TBTIPOMAT AS A "
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Código", "Descrição", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListViewTipoMaterial
        MontaDadosLVTeste "S", vListViewTipoMaterial
        
        If checaFiltro = True Then
            PersonaColLVTeste 2, "N", "N", "", "S", "N", "N", "E", vListViewTipoMaterial
        End If
        If vListViewTipoMaterial.ListItems.Count > 0 Then ajusta_LVTeste vListViewTipoMaterial
    End If
    
    
'--CLIENTES
    If QualLV = 1 Then
        Set vListviewClientes = vListViewPrincipal
        Set chamaForm = New frmClientes
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 2
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbclifor", "", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then SqlLV = "SELECT A.CODCLIFOR,A.NOME,A.ENDERECO,A.CEP,A.BAIRRO,A.CIDADE,A.UF,A.ATIVO FROM TBCLIFOR AS A ORDER BY A.CODCLIFOR "
            If FiltroGeral = "Ativos" Then SqlLV = "Select a.codclifor,a.nome,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo  from tbclifor as a where a.ativo='S' Order by a.codclifor "
            If FiltroGeral = "Não ativos" Then SqlLV = "Select a.codclifor,a.nome,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo  from tbclifor as a where a.ativo<>'S' Order by a.codclifor "
        Else
            If frmPesqGeralTeste2.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Código", "Nome", "Endereço", "CEP", "Bairro", "Cidade", "UF", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewClientes
        MontaDadosLVTeste "S", vListviewClientes
        If checaFiltro = True Then
            PersonaColLVTeste 7, "N", "P", "", "S", "N", "N", "E", vListviewClientes
        End If
        If vListviewClientes.ListItems.Count > 0 Then ajusta_LVTeste vListviewClientes
    End If

'--PARADAS
    If QualLV = 2 Then
        Set vListViewParadas = vListViewPrincipal
        Set chamaForm = New frmAtividades
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 2
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbparadas", "", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then SqlLV = "SELECT A.IDPARADA,A.TIPO,A.CODIGO,A.NMPARADA,A.DESCRICAO,A.ATIVO FROM TBPARADAS AS A WHERE A.IDPARADA IS NOT NULL"
            If FiltroGeral = "Ativos" Then SqlLV = "Select a.idparada,a.tipo,a.codigo,a.nmparada,a.descricao,a.ativo from tbparadas as a where a.ativo = 'S'"
            If FiltroGeral = "Não ativos" Then SqlLV = "Select a.idparada,a.tipo,a.codigo,a.nmparada,a.descricao,a.ativo from tbparadas as a where a.ativo <> 'S'"
        Else
            If frmPesqGeralTeste2.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "ID", "Tipo", "Código", "Nome", "Descrição", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListViewParadas
        MontaDadosLVTeste "S", vListViewParadas
        If checaFiltro = True Then
            PersonaColLVTeste 5, "N", "P", "", "S", "N", "N", "E", vListViewParadas
        End If
        If vListViewParadas.ListItems.Count > 0 Then ajusta_LVTeste vListViewParadas
    End If

'--TRANSPORTADORAS
    If QualLV = 3 Then
        Set vListviewTransportadoras = vListViewPrincipal
        Set chamaForm = New frmTransportes
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "ttra", "", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then SqlLV = "SELECT A.CODTRA,A.NOME,A.CGC,A.INSCRESTADUAL,A.RUA+','+A.NUMERO AS ENDERECO,A.CEP,A.BAIRRO,A.CIDADE,A.CODETD,A.INATIVO FROM " & vBancoTotvs & ".DBO.TTRA AS A WHERE A.INATIVO = 0 OR A.INATIVO IS NULL"
        Else
            If frmPesqGeralTeste2.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Código", "Nome", "CNPJ", "IE", "Endereço", "CEP", "Bairro", "Cidade", "UF", "Ativo", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewTransportadoras
        MontaDadosLVTeste "N", vListviewTransportadoras
        If checaFiltro = True Then
            PersonaColLVTeste 9, "N", "P", "", "S", "N", "N", "E", vListviewTransportadoras
        End If
        If vListviewTransportadoras.ListItems.Count > 0 Then ajusta_LVTeste vListviewTransportadoras
    End If
    
'--FÓRMULAS PRODUTOS
    If QualLV = 4 Then
        Set vListviewFormulaPRD = vListViewPrincipal
        Set chamaForm = New frmMaterial
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "TPRD", "tbMateriais", "ttb2", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                'SqlLV = "select a.idprd,a.CODIGOPRD,a.NOMEFANTASIA,a.codtb2fat,c.descricao,b.formula,b.forpint from " & vBancoTotvs & ".dbo.TPRD as a left join " & sDatabaseName & ".dbo.tbMateriais as b on a.IDPRD = b.IDPRD left join " & vBancoTotvs & ".dbo.ttb2 as c on a.CODTB2FAT = c.CODTB2FAT and c.CODCOLIGADA = " & vCodcoligada & " where a.CODIGOPRD like '%%' and a.CODCOLIGADA = " & vCodcoligada
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrL
                SqlLV = SqlLV & " A.IDPRD, " & vbCrLf
                SqlLV = SqlLV & " A.CODIGOPRD, " & vbCrLf
                SqlLV = SqlLV & " A.NOMEFANTASIA, " & vbCrLf
                SqlLV = SqlLV & " A.CODTB2FAT, " & vbCrLf
                SqlLV = SqlLV & " C.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " B.FORMULA, " & vbCrLf
                SqlLV = SqlLV & " B.FORPINT " & vbCrLf
                SqlLV = SqlLV & "FROM " & vBancoTotvs & ".DBO.TPRD AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & sDatabaseName & ".DBO.TBMATERIAIS AS B ON A.IDPRD = B.IDPRD " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vBancoTotvs & ".DBO.TTB2 AS C ON " & vbCrLf
                SqlLV = SqlLV & " A.CODTB2FAT = C.CODTB2FAT AND C.CODCOLIGADA = " & vCodcoligada & " WHERE A.IDPRD IS NOT NULL"
            End If
        Else
            If frmPesqGeralTeste2.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "ID", "Código", "Descrição", "Cod Tipo", "Tipo Material", "Fórmula PESO", "Fórmula PINTURA", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewFormulaPRD
        MontaDadosLVTeste "N", vListviewFormulaPRD
        If checaFiltro = True Then
            'PersonaColLVTeste 6, "N", "N", "", "S", "N", "N", "E", vListviewFormulaPRD
        End If
        If vListviewFormulaPRD.ListItems.Count > 0 Then ajusta_LVTeste vListviewFormulaPRD
    End If
    
'--ORÇAMENTOS
    If QualLV = 5 Then
        Set vListviewComercial = vListViewPrincipal
        Set chamaForm = New frmFO
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        Permissao
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbclifor", "tbfo", "tbcontatos", "tbfce", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then 'SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo,CASE WHEN D.status = 0 THEN 'ANDAMENTO' WHEN D.status = 1 THEN 'CONCLUIDA' WHEN D.status = 2 THEN 'PARALIZADA' END AS STATUS from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato left join tbFCE as d on b.fce = d.fce"
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " FO.CODFO, " & vbCrLf
                SqlLV = SqlLV & " FO.NOME, " & vbCrLf
                SqlLV = SqlLV & " FO.PEDIDO, " & vbCrLf
                SqlLV = SqlLV & " FO.NOME_CONTATO, " & vbCrLf
                SqlLV = SqlLV & " FO.TELEFONE, " & vbCrLf
                SqlLV = SqlLV & " FO.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " FO.DATAFO, " & vbCrLf
                SqlLV = SqlLV & " FO.DATADEVCP, " & vbCrLf
                SqlLV = SqlLV & " FO.PROPOSTA, " & vbCrLf
                SqlLV = SqlLV & " FO.QUANTIDADE, " & vbCrLf
                SqlLV = SqlLV & " FO.VALORUNIT, " & vbCrLf
                SqlLV = SqlLV & " FO.VALORTOTAL, " & vbCrLf
                SqlLV = SqlLV & " FO.PEDIDO1, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FO.FCE IS NULL THEN 0 ELSE FO.FCE END AS FCE, " & vbCrLf
                SqlLV = SqlLV & " FO.STATUSFO, " & vbCrLf
                SqlLV = SqlLV & " FO.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " FO.STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS [TIPO FCE] " & vbCrLf
                SqlLV = SqlLV & " FROM ( " & vbCrLf
                SqlLV = SqlLV & " SELECT " & vbCrLf
                SqlLV = SqlLV & "     B.CODFO AS CODFO, " & vbCrLf
                SqlLV = SqlLV & "     A.NOME AS NOME, " & vbCrLf
                SqlLV = SqlLV & "     B.PEDIDO AS PEDIDO, " & vbCrLf
                SqlLV = SqlLV & "     C.NOME AS NOME_CONTATO, " & vbCrLf
                SqlLV = SqlLV & "     C.TELEFONE AS TELEFONE, " & vbCrLf
                SqlLV = SqlLV & "     B.DESCRICAO AS DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & "     B.DATAFO AS DATAFO, " & vbCrLf
                SqlLV = SqlLV & "     B.DATADEVCP AS DATADEVCP, " & vbCrLf
                SqlLV = SqlLV & "     B.PROPOSTA AS PROPOSTA, " & vbCrLf
                SqlLV = SqlLV & "     B.QUANTIDADE AS QUANTIDADE, " & vbCrLf
                SqlLV = SqlLV & "     B.VALORUNIT AS VALORUNIT, " & vbCrLf
                SqlLV = SqlLV & "     (B.QUANTIDADE*B.VALORUNIT) AS VALORTOTAL, " & vbCrLf
                SqlLV = SqlLV & "     B.PEDIDO AS PEDIDO1, " & vbCrLf
                SqlLV = SqlLV & "     B.FCE AS FCE, " & vbCrLf
                SqlLV = SqlLV & "     B.STATUSFO AS STATUSFO, " & vbCrLf
                SqlLV = SqlLV & "     B.ATIVO ATIVO, " & vbCrLf
                SqlLV = SqlLV & "     CASE WHEN D.STATUS = 0 THEN 'ANDAMENTO' WHEN D.STATUS = 1 THEN 'CONCLUIDA' WHEN D.STATUS = 2 THEN 'PARALIZADA' END AS STATUS " & vbCrLf
                SqlLV = SqlLV & " FROM " & vbCrLf
                SqlLV = SqlLV & " TBCLIFOR AS A " & vbCrLf
                SqlLV = SqlLV & " INNER JOIN TBFO  AS B ON A.CODCLIFOR=B.CODCLIFOR " & vbCrLf
                SqlLV = SqlLV & " LEFT JOIN TBCONTATOS AS C ON B.CODCLIFOR = C.CODCLIFOR AND B.CODCONTATO = C.CODCONTATO " & vbCrLf
                SqlLV = SqlLV & " LEFT JOIN TBFCE AS D ON B.FCE = D.FCE " & vbCrLf
                SqlLV = SqlLV & ") AS FO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON " & vbCrLf
                SqlLV = SqlLV & "FO.FCE = FILTRO.FCE WHERE FO.CODFO IS NOT NULL ORDER BY FO.CODFO DESC"
            End If
            If FiltroGeral = "Ativos" Then SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo,CASE WHEN D.status = 0 THEN 'ANDAMENTO' WHEN D.status = 1 THEN 'CONCLUIDA' WHEN D.status = 2 THEN 'PARALIZADA' END AS STATUS from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato left join tbFCE as d on b.fce = d.fce where b.ativo = 'S'"
            If FiltroGeral = "Não ativos" Then SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo,CASE WHEN D.status = 0 THEN 'ANDAMENTO' WHEN D.status = 1 THEN 'CONCLUIDA' WHEN D.status = 2 THEN 'PARALIZADA' END AS STATUS from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato left join tbFCE as d on b.fce = d.fce where b.ativo <> 'S'"
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        
        MontaCabLV "FO", "Empresa", "Coleta nº", "Contato", "Fone", "Descrição", "Data Abertura", "Dev. CP", "Proposta nº", "Quant.", "Valor Unit", "Valor Total", "Pedido nº", "FCE nº", "Status FO", "Ativo", "Status FCE", "Tipo FCE", "", "", ""
        MontaCabecalhoLVTeste vListviewComercial
        MontaDadosLVTeste "N", vListviewComercial
        If checaFiltro = True Then
            PersonaColLVTeste 14, "N", "N", "", "S", "N", "N", "E", vListviewComercial
            PersonaColLVTeste 13, "S", "S", "", "N", "N", "N", "D", vListviewComercial
            PersonaColLVTeste 15, "N", "N", "", "S", "N", "N", "E", vListviewComercial
            PersonaColLVTeste 16, "N", "N", "", "S", "N", "N", "E", vListviewComercial
            PersonaColLVTeste 17, "N", "P", "", "N", "N", "N", "E", vListviewComercial 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
        End If
        If vListviewComercial.ListItems.Count > 0 Then ajusta_LVTeste vListviewComercial
    End If
    
'--FCE - Ficha de Controle de Encomenda
    If QualLV = 6 Then
        Set vListviewFCE = vListViewPrincipal
        Set chamaForm = New frmFCECons
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 3
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbfo", "tbfce", "tbclifor", "tbcontatos", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " B.DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " C.NOME[CLIENTE], " & vbCrLf
                SqlLV = SqlLV & " D.NOME[CONTATO], " & vbCrLf
                SqlLV = SqlLV & " D.TELEFONE, " & vbCrLf
                SqlLV = SqlLV & " B.DATAENTREGA, " & vbCrLf
                SqlLV = SqlLV & " B.PINTURA, " & vbCrLf
                SqlLV = SqlLV & " B.TRANSPORTE, " & vbCrLf
                SqlLV = SqlLV & " B.MATERIAPRIMA, " & vbCrLf
                SqlLV = SqlLV & " B.FABRICACAO, " & vbCrLf
                SqlLV = SqlLV & " B.REPARO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN B.STATUS = 0 THEN 'ANDAMENTO' WHEN B.STATUS = 1 THEN 'CONCLUIDA' WHEN B.STATUS = 2 THEN 'PARALIZADA' END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE FROM TBFO AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON B.FCE = A.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBCLIFOR AS C ON A.CODCLIFOR=C.CODCLIFOR " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBCONTATOS AS D ON " & vbCrLf
                SqlLV = SqlLV & " C.CODCLIFOR = D.CODCLIFOR AND " & vbCrLf
                SqlLV = SqlLV & " D.CODCONTATO = A.CODCONTATO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
                SqlLV = SqlLV & " WHERE B.FCE IS NOT NULL ORDER BY B.FCE DESC"
            End If
            If FiltroGeral = "Ativos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " B.DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " C.NOME[CLIENTE], " & vbCrLf
                SqlLV = SqlLV & " D.NOME[CONTATO], " & vbCrLf
                SqlLV = SqlLV & " D.TELEFONE, " & vbCrLf
                SqlLV = SqlLV & " B.DATAENTREGA, " & vbCrLf
                SqlLV = SqlLV & " B.PINTURA, " & vbCrLf
                SqlLV = SqlLV & " B.TRANSPORTE, " & vbCrLf
                SqlLV = SqlLV & " B.MATERIAPRIMA, " & vbCrLf
                SqlLV = SqlLV & " B.FABRICACAO, " & vbCrLf
                SqlLV = SqlLV & " B.REPARO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN B.STATUS = 0 THEN 'ANDAMENTO' WHEN B.STATUS = 1 THEN 'CONCLUIDA' WHEN B.STATUS = 2 THEN 'PARALIZADA' END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBFO AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON B.FCE = A.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBCLIFOR AS C ON A.CODCLIFOR=C.CODCLIFOR " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBCONTATOS AS D ON " & vbCrLf
                SqlLV = SqlLV & " C.CODCLIFOR = D.CODCLIFOR AND " & vbCrLf
                SqlLV = SqlLV & " D.CODCONTATO = A.CODCONTATO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE  WHERE A.ATIVO = 'S' " & vbCrLf
                SqlLV = SqlLV & "ORDER BY B.FCE DESC"
            End If
            If FiltroGeral = "Não ativos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " B.DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " C.NOME[CLIENTE], " & vbCrLf
                SqlLV = SqlLV & " D.NOME[CONTATO], " & vbCrLf
                SqlLV = SqlLV & " D.TELEFONE, " & vbCrLf
                SqlLV = SqlLV & " B.DATAENTREGA, " & vbCrLf
                SqlLV = SqlLV & " B.PINTURA, " & vbCrLf
                SqlLV = SqlLV & " B.TRANSPORTE, " & vbCrLf
                SqlLV = SqlLV & " B.MATERIAPRIMA, " & vbCrLf
                SqlLV = SqlLV & " B.FABRICACAO, " & vbCrLf
                SqlLV = SqlLV & " B.REPARO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN B.STATUS = 0 THEN 'ANDAMENTO' WHEN B.STATUS = 1 THEN 'CONCLUIDA' WHEN B.STATUS = 2 THEN 'PARALIZADA' END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBFO AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON B.FCE = A.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBCLIFOR AS C ON A.CODCLIFOR=C.CODCLIFOR " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBCONTATOS AS D ON " & vbCrLf
                SqlLV = SqlLV & " C.CODCLIFOR = D.CODCLIFOR AND " & vbCrLf
                SqlLV = SqlLV & " D.CODCONTATO = A.CODCONTATO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE  WHERE A.ATIVO <> 'S' " & vbCrLf
                SqlLV = SqlLV & "ORDER BY B.FCE DESC"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Data abertura", "FCE", "Cliente", "Contato", "Fone", "Data entrega", "Pintura", "Transporte", "Matéria-prima", "Fabricação", "Reparo", "Ativo", "Status FCE", "Tipo FCE", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewFCE
        MontaDadosLVTeste "N", vListviewFCE
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "D", vListviewFCE
            PersonaColLVTeste 11, "N", "N", "", "S", "N", "N", "E", vListviewFCE
            PersonaColLVTeste 12, "N", "N", "", "S", "N", "N", "E", vListviewFCE
            PersonaColLVTeste 13, "N", "P", "", "N", "N", "N", "E", vListviewFCE 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
        End If
        If vListviewFCE.ListItems.Count > 0 Then ajusta_LVTeste vListviewFCE
    End If
    
'--DESENHOS
    If QualLV = 7 Then
        Set vListviewDesenhos = vListViewPrincipal
        Set chamaForm = New frmDesenhos
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbDesenhos", "tbProjetos", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.IDDESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.DESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.REVISAO, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
                SqlLV = SqlLV & " A.TIPO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBDESENHOS AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "WHERE A.IDDESENHO IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "ORDER BY " & vbCrLf
                SqlLV = SqlLV & " B.FCE DESC,B.PROJETO"
            End If
            If FiltroGeral = "Ativos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.IDDESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.DESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.REVISAO, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
                SqlLV = SqlLV & " A.TIPO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBDESENHOS AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "WHERE A.CODCOLIGADA = '" & vCodcoligada & "' AND A.ATIVO = 'S' " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "ORDER BY " & vbCrLf
                SqlLV = SqlLV & " B.FCE DESC,B.PROJETO"
            End If
            If FiltroGeral = "Não ativos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.IDDESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.DESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.REVISAO, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
                SqlLV = SqlLV & " A.TIPO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBDESENHOS AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "WHERE A.CODCOLIGADA = '" & vCodcoligada & "' AND A.ATIVO = 'N' " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "ORDER BY " & vbCrLf
                SqlLV = SqlLV & " B.FCE DESC,B.PROJETO"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Identificador", "Desenho", "Rev.", "FCE", "Projeto", "Data Cadastro", "Tipo", "Ativo", "Tipo FCE", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewDesenhos
        MontaDadosLVTeste "S", vListviewDesenhos
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "N", "", "N", "N", "N", "E", vListviewDesenhos
            PersonaColLVTeste 7, "N", "N", "", "S", "N", "N", "E", vListviewDesenhos
            PersonaColLVTeste 8, "N", "P", "", "N", "N", "N", "E", vListviewDesenhos 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
        End If
        If vListviewDesenhos.ListItems.Count > 0 Then ajusta_LVTeste vListviewDesenhos
    End If
    
'--LM - LISTA DE MATERIAIS
    If QualLV = 8 Then
        Set vListviewLM = vListViewPrincipal
        Set chamaForm = New frmLM
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 2 'Com quantas colunas que a varglobal irá trabalhar
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tblm", "tbfce", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " RIGHT('0000' + RTRIM(A.FCE),4) AS FCE, " & vbCrLf
                SqlLV = SqlLV & " A.CODLM AS LM, " & vbCrLf
                SqlLV = SqlLV & " A.DATAABERTURA AS DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " A.DESCRICAO AS DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO AS ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 1 THEN " & vbCrLf
                SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 2 THEN " & vbCrLf
                SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
                SqlLV = SqlLV & "         A.TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "     ELSE " & vbCrLf
                SqlLV = SqlLV & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
                SqlLV = SqlLV & " END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBLM AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON A.FCE = B.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON A.FCE = FILTRO.FCE " & vbCrLf
                SqlLV = SqlLV & "WHERE A.CODLM IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.FCE DESC, A.CODLM DESC"
            End If
            If FiltroGeral = "Ativos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " RIGHT('0000' + RTRIM(A.FCE),4) AS FCE, " & vbCrLf
                SqlLV = SqlLV & " A.CODLM AS LM, " & vbCrLf
                SqlLV = SqlLV & " A.DATAABERTURA AS DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " A.DESCRICAO AS DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO AS ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 1 THEN " & vbCrLf
                SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 2 THEN " & vbCrLf
                SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
                SqlLV = SqlLV & "         A.TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "     ELSE " & vbCrLf
                SqlLV = SqlLV & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
                SqlLV = SqlLV & " END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBLM AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON A.FCE = B.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON A.FCE = FILTRO.FCE WHERE A.ATIVO = 'S' " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.FCE DESC, A.CODLM DESC"
            End If
            If FiltroGeral = "Não ativos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " RIGHT('0000' + RTRIM(A.FCE),4) AS FCE, " & vbCrLf
                SqlLV = SqlLV & " A.CODLM AS LM, " & vbCrLf
                SqlLV = SqlLV & " A.DATAABERTURA AS DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " A.DESCRICAO AS DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO AS ATIVO, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 1 THEN " & vbCrLf
                SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN B.STATUS = 2 THEN " & vbCrLf
                SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
                SqlLV = SqlLV & "         A.TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "     ELSE " & vbCrLf
                SqlLV = SqlLV & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
                SqlLV = SqlLV & " END AS TIPO_FCE " & vbCrLf
                SqlLV = SqlLV & "FROM TBLM AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON A.FCE = B.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON A.FCE = FILTRO.FCE WHERE A.ATIVO <> 'S' " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.FCE DESC, A.CODLM DESC"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "FCE", "LM", "Data Abertura", "Descrição", "Ativo", "Status FCE", "Tipo FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewLM
        MontaDadosLVTeste "N", vListviewLM
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "S", "", "N", "S", "N", "D", vListviewLM
            PersonaColLVTeste 4, "N", "N", "", "S", "N", "N", "E", vListviewLM
            PersonaColLVTeste 5, "N", "N", "", "S", "N", "N", "E", vListviewLM
            PersonaColLVTeste 6, "N", "P", "", "N", "N", "N", "E", vListviewLM  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
        End If
        If vListviewLM.ListItems.Count > 0 Then ajusta_LVTeste vListviewLM
    End If
    
'--MP - Métodos e Processos
    If QualLV = 9 Then
        Set vListviewMP = vListViewPrincipal
        Set chamaForm = New frmMPCompleto
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1 'Com quantas colunas que a varglobal irá trabalhar
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbMP", "tbProjetos", "tbMPItens", "tbitemlm", "tbdesenhos", "tbos", "tbretrabalho", "tcfce", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.IDPROGRAMACAO, " & vbCrLf
                SqlLV = SqlLV & " C.IDOS, " & vbCrLf
                SqlLV = SqlLV & " F.REVISAO, " & vbCrLf
                SqlLV = SqlLV & " A.DATAPROGRAMACAO, " & vbCrLf
                SqlLV = SqlLV & " B.FCE, " & vbCrLf
                SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " A.RESPONSAVEL, " & vbCrLf
                SqlLV = SqlLV & " MIN(E.DESENHO) AS DESENHO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
                SqlLV = SqlLV & " MAX(G.IDRETRABALHO) AS RETRABALHO, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.STATUS = 1 THEN 'Planejamento' " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.STATUS > 1 AND A.STATUS < 3 THEN 'Produção' " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.STATUS = 3 THEN 'Expedição' " & vbCrLf
                SqlLV = SqlLV & "     ELSE 'Planejamento' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN H.STATUS = 0 THEN 'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN H.STATUS = 1 THEN 'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN H.STATUS = 2 THEN 'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS_FCE, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN F.TIPOOS = 0 THEN 'Fabricação' WHEN F.TIPOOS = 1 THEN 'Manutenção' WHEN F.TIPOOS = 2 THEN 'Usinagem' ELSE 'Fabricação' END AS TIPO " & vbCrLf
                SqlLV = SqlLV & "FROM TBMP AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBPROJETOS AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBMPITENS AS C ON " & vbCrLf
                SqlLV = SqlLV & " A.IDPROGRAMACAO = C.IDPROGRAMACAO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBITEMLM AS D ON " & vbCrLf
                SqlLV = SqlLV & " SUBSTRING(C.DESENHOS,1,2) = D.CODLM AND " & vbCrLf
                SqlLV = SqlLV & " REPLACE(SUBSTRING(C.DESENHOS,3,4),';','') = D.CODSEQ AND " & vbCrLf
                SqlLV = SqlLV & " B.FCE = D.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBDESENHOS AS E ON " & vbCrLf
                SqlLV = SqlLV & " D.CODIGODES = E.IDDESENHO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBOS AS F ON " & vbCrLf
                SqlLV = SqlLV & " C.IDOS = F.IDOS AND " & vbCrLf
                SqlLV = SqlLV & " C.REVISAOOS = F.REVISAO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBRETRABALHO AS G ON " & vbCrLf
                SqlLV = SqlLV & " A.IDPROGRAMACAO = G.IDPROGRAMACAO " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS H ON " & vbCrLf
                SqlLV = SqlLV & " B.FCE = H.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
                SqlLV = SqlLV & "     COALESCE( " & vbCrLf
                SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
                SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
                SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
                SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
                SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
                SqlLV = SqlLV & " FROM TBLM AS C " & vbCrLf
                SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
                SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO = 'S' " & vbCrLf
                SqlLV = SqlLV & "GROUP BY A.IDPROGRAMACAO,C.IDOS,F.REVISAO,A.DATAPROGRAMACAO,B.FCE,B.PROJETO,A.RESPONSAVEL,A.ATIVO,A.STATUS,H.STATUS,FILTRO.TIPO,F.TIPOOS " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.IDPROGRAMACAO DESC"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        QtdColReal = 0
        MontaCabLV "Planejamento", "OS nº", "Rev.", "Data", "FCE", "Projeto", "Responsável", "Desenho", "Ativo", "Retrabalho", "Status", "Status FCE", "Tipo FCE", "Tipo OS", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewMP
        MontaDadosLVTeste "S", vListviewMP
        If checaFiltro = True Then
            PersonaColLVTeste 1, "N", "N", "", "N", "S", "N", "E", vListviewMP
            PersonaColLVTeste 8, "N", "N", "", "S", "N", "N", "E", vListviewMP
            PersonaColLVTeste 9, "S", "N", "", "N", "N", "N", "E", vListviewMP
            PersonaColLVTeste 10, "N", "S", "", "N", "N", "N", "E", vListviewMP
            PersonaColLVTeste 11, "N", "N", "", "S", "N", "N", "E", vListviewMP
            PersonaColLVTeste 12, "N", "P", "", "N", "N", "N", "E", vListviewMP  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
        End If
        If vListviewMP.ListItems.Count > 0 Then ajusta_LVTeste vListviewMP
    End If
    
'-- CONTROLE DE DESENHOS
    If QualLV = 10 Then
        Set vListviewControleDesenhos = vListViewPrincipal
        Set chamaForm = New frmCD
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            frmFiltro.frmPeriodo.Visible = True
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbcd", "tbdesenhos", "tbprojetos", "", "", "", "", "", "", ""
            
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.IDCD, " & vbCrLf
                SqlLV = SqlLV & " CAST(C.FCE AS VARCHAR(4)) + ' - ' + C.PROJETO AS FCE, " & vbCrLf
                SqlLV = SqlLV & " B.DESENHO, " & vbCrLf
                SqlLV = SqlLV & " B.REVISAO, " & vbCrLf
                SqlLV = SqlLV & " A.QUANTIDADE, " & vbCrLf
                SqlLV = SqlLV & " A.PESOUNIT, " & vbCrLf
                SqlLV = SqlLV & " (A.QUANTIDADE*A.PESOUNIT) AS PESOTOTAL, " & vbCrLf
                SqlLV = SqlLV & " A.DATARECEBIDO,A.PTEMPO + ' ' + A.PUNIDADE, " & vbCrLf
                SqlLV = SqlLV & " A.USUARIO, " & vbCrLf
                SqlLV = SqlLV & " A.DATAINI, " & vbCrLf
                SqlLV = SqlLV & " A.DATAFIM, " & vbCrLf
                SqlLV = SqlLV & " A.CROQUI, " & vbCrLf
                SqlLV = SqlLV & " A.STATUS, " & vbCrLf
                SqlLV = SqlLV & " A.OBSERVACAO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO,A.DETALHISTA " & vbCrLf
                SqlLV = SqlLV & "FROM TBCD AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBDESENHOS AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.IDDESENHO = B.IDDESENHO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS C ON " & vbCrLf
                SqlLV = SqlLV & " B.CODPROJETO = C.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " A.IDCD IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.IDCD DESC"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Identificador", "FCE", "Desenho", "Rev.", "Quant.", "Peso Unit.", "Peso Total", "Recebido", "Previsão Det.", "Usuário", "Data inicio", "Data fim", "Croqui", "Status", "Observação", "Ativo", "Detalhista", "", "", "", ""
        MontaCabecalhoLVTeste vListviewControleDesenhos
        MontaDadosLVTeste "S", vListviewControleDesenhos
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "N", "", "N", "N", "N", "E", vListviewControleDesenhos
            PersonaColLVTeste 4, "N", "N", "", "N", "N", "N", "D", vListviewControleDesenhos
            PersonaColLVTeste 5, "N", "N", "", "N", "N", "S", "D", vListviewControleDesenhos
            PersonaColLVTeste 6, "N", "N", "", "N", "N", "S", "D", vListviewControleDesenhos
            
            PersonaColLVTeste 13, "N", "N", "", "S", "N", "N", "E", vListviewControleDesenhos
            PersonaColLVTeste 15, "N", "P", "", "S", "N", "N", "E", vListviewControleDesenhos
        End If
        If vListviewControleDesenhos.ListItems.Count > 0 Then ajusta_LVTeste vListviewControleDesenhos
    End If
    
'--FÓRMULA CENTRO DE CUSTO
    If QualLV = 11 Then
        Set vListviewFormulaCC = vListViewPrincipal
        Set chamaForm = New frmFormulaCC
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "GCCUSTO", "tbFormula", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODREDUZIDO, " & vbCrLf
                SqlLV = SqlLV & " A.NOME, " & vbCrLf
                SqlLV = SqlLV & " 'FORMULA' = " & vbCrLf
                SqlLV = SqlLV & "     CASE " & vbCrLf
                SqlLV = SqlLV & "         WHEN MAX(B.NMFORM) IS NULL THEN " & vbCrLf
                SqlLV = SqlLV & "             '-' " & vbCrLf
                SqlLV = SqlLV & "         ELSE " & vbCrLf
                SqlLV = SqlLV & "             'COM FORMULA' " & vbCrLf
                SqlLV = SqlLV & "     END " & vbCrLf
                SqlLV = SqlLV & "FROM " & vBancoTotvs & ".DBO.GCCUSTO AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN  " & sDatabaseName & ".DBO.TBFORMULA AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.CODREDUZIDO = B.CODREDUZIDO COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " A.CODCOLIGADA = '" & vCodcoligada & "' AND " & vbCrLf
                SqlLV = SqlLV & " ATIVO  = 'T' OR " & vbCrLf
                SqlLV = SqlLV & "  " & vbCrLf
                SqlLV = SqlLV & " A.CODCOLIGADA = '" & vCodcoligada & "' AND " & vbCrLf
                SqlLV = SqlLV & " ATIVO  = 'T' AND " & vbCrLf
                SqlLV = SqlLV & " SUBSTRING(A.CODREDUZIDO,1,4) IN('1000','3000','7000','5000','6000','9001','7000','4000') " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & "GROUP BY A.ID,A.CODREDUZIDO,A.NOME " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.CODREDUZIDO"
            End If
'            SqlLV = "select a.CODREDUZIDO,a.NOME, 'formula' = case when max(b.nmform) IS NULL then '-' else 'com formula' end from " & vBancoTotvs & ".dbo.GCCUSTO as a left join " & sDatabaseName & ".dbo.tbFormula as b " & _
'            "on a.CODREDUZIDO = b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS where a.codcoligada = '" & vCodcoligada & "' and (ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '1000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '3000' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '7000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '5000' or " & _
'            "ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '6000' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '9001' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '7000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '4000' or ativo  = 'T') group by a.ID,a.CODREDUZIDO,a.NOME order by a.CODREDUZIDO"
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        QtdColReal = 0
        MontaCabLV "Centro de Custo", "Nome Centro de Custo", "Fórmula", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewFormulaCC
        MontaDadosLVTeste "S", vListviewFormulaCC
        If checaFiltro = True Then
            'PersonaColLVTeste 3, "N", "N", "", "N", "N", "N", "D", vListviewFormulaCC
            'PersonaColLVTeste 4, "N", "N", "", "S", "N", "N", "E", vListviewFormulaCC
        End If
        If vListviewFormulaCC.ListItems.Count > 0 Then ajusta_LVTeste vListviewFormulaCC
    End If
    
'-- QUALIDADE - RNCF (Registro de Não Conformidade de Fabricação)
    If QualLV = 12 Then
        Set vListviewRNCF = vListViewPrincipal
        Set chamaForm = New frmRNCF
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbComunicacaoDesvio", "tbMPItens", "tbMP", "tbProjetos", "tbRNC", "tbRetrabalho", "", "", "", ""
            
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.IDCD, " & vbCrLf
                SqlLV = SqlLV & " A.DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " A.RESPONSAVEL, " & vbCrLf
                SqlLV = SqlLV & " A.IDOS, " & vbCrLf
                SqlLV = SqlLV & " D.FCE, " & vbCrLf
                SqlLV = SqlLV & " SUBSTRING(D.PROJETO,1,20) AS PROJETO, " & vbCrLf
                SqlLV = SqlLV & " SUBSTRING(A.OBSERVACAO,1,100) AS OBSERVACAO, " & vbCrLf
                SqlLV = SqlLV & " A.STATUS, " & vbCrLf
                SqlLV = SqlLV & " E.IDRNC, " & vbCrLf
                SqlLV = SqlLV & " E.DATACONCLUSAO, " & vbCrLf
                SqlLV = SqlLV & " E.GEROURETRABALHO, " & vbCrLf
                SqlLV = SqlLV & " H.IDRETRABALHO, " & vbCrLf
                SqlLV = SqlLV & " E.DATAFECHAMENTO " & vbCrLf
                SqlLV = SqlLV & "FROM TBCOMUNICACAODESVIO AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBMPITENS AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.IDOS = B.IDOS " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBMP AS C ON " & vbCrLf
                SqlLV = SqlLV & " B.IDPROGRAMACAO = C.IDPROGRAMACAO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS D ON " & vbCrLf
                SqlLV = SqlLV & " C.CODPROJETO = D.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBRNC AS E ON " & vbCrLf
                SqlLV = SqlLV & " A.IDCD = E.IDCD " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBRETRABALHO AS H ON " & vbCrLf
                SqlLV = SqlLV & " A.IDCD = H.IDCD " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " A.IDCD IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & "GROUP BY " & vbCrLf
                SqlLV = SqlLV & " A.IDCD, " & vbCrLf
                SqlLV = SqlLV & " A.DATAABERTURA, " & vbCrLf
                SqlLV = SqlLV & " A.RESPONSAVEL, " & vbCrLf
                SqlLV = SqlLV & " A.IDOS, " & vbCrLf
                SqlLV = SqlLV & " D.FCE, " & vbCrLf
                SqlLV = SqlLV & " SUBSTRING(D.PROJETO,1,20), " & vbCrLf
                SqlLV = SqlLV & " SUBSTRING(A.OBSERVACAO,1,100), " & vbCrLf
                SqlLV = SqlLV & " A.STATUS, " & vbCrLf
                SqlLV = SqlLV & " E.IDRNC, " & vbCrLf
                SqlLV = SqlLV & " E.DATACONCLUSAO, " & vbCrLf
                SqlLV = SqlLV & " E.GEROURETRABALHO, " & vbCrLf
                SqlLV = SqlLV & " H.IDRETRABALHO, " & vbCrLf
                SqlLV = SqlLV & " E.DATAFECHAMENTO " & vbCrLf
                SqlLV = SqlLV & "ORDER BY " & vbCrLf
                SqlLV = SqlLV & " A.IDCD DESC"
            End If
                'SqlLV = "select top " & LimiteLinhas & " a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
                '                                  "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd where a.idcd >= 1 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "CD nº", "Data Abertura", "Responsável", "OS nº", "FCE", "Projeto", "Observação", "Status", "RNC nº", "Data Conclusão", "Retrabalho", "Retrabalho nº", "Data Fechamento", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewRNCF
        MontaDadosLVTeste "S", vListviewRNCF
        If checaFiltro = True Then
            PersonaColLVTeste 3, "N", "N", "", "N", "S", "N", "E", vListviewRNCF
            PersonaColLVTeste 7, "S", "S", "", "S", "N", "N", "E", vListviewRNCF
            PersonaColLVTeste 8, "S", "N", "", "N", "S", "N", "E", vListviewRNCF
            PersonaColLVTeste 10, "N", "P", "", "S", "N", "N", "E", vListviewRNCF
            PersonaColLVTeste 11, "S", "S", "", "N", "S", "N", "E", vListviewRNCF
        End If
        If vListviewRNCF.ListItems.Count > 0 Then ajusta_LVTeste vListviewRNCF
    End If
    
'-- USUÁRIOS
    If QualLV = 13 Then
        Set vListviewUsuarios = vListViewPrincipal
        Set chamaForm = New frmUsuarios
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbusuarios", "tbgrupo", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODIGO, " & vbCrLf
                SqlLV = SqlLV & " A.NOME, " & vbCrLf
                SqlLV = SqlLV & " B.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO " & vbCrLf
                SqlLV = SqlLV & "FROM TBUSUARIOS AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBGRUPO AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.CODGRUPO = B.CODIGO " & vbCrLf
                SqlLV = SqlLV & "WHERE B.CODCOLIGADA = " & vCodcoligada
            End If
            'SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where b.codcoligada = " & vCodcoligada
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Código", "Nome do usuário", "Grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewUsuarios
        MontaDadosLVTeste "S", vListviewUsuarios
        If checaFiltro = True Then
            PersonaColLVTeste 3, "N", "P", "", "S", "N", "N", "E", vListviewUsuarios
        End If
        If vListviewUsuarios.ListItems.Count > 0 Then ajusta_LVTeste vListviewUsuarios
    End If
    
'-- GRUPOS
    If QualLV = 14 Then
        Set vListviewGrupos = vListViewPrincipal
        Set chamaForm = New frmGrupos
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            
            carregaTABS "tbGrupo", "", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODIGO, " & vbCrLf
                SqlLV = SqlLV & " A.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO " & vbCrLf
                SqlLV = SqlLV & "FROM TBGRUPO AS A " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " CODCOLIGADA = " & vCodcoligada
                'SqlLV = "select a.codigo,a.descricao,a.ativo from tbgrupo as a"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Código", "Nome do grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewGrupos
        MontaDadosLVTeste "S", vListviewGrupos
        If checaFiltro = True Then
            PersonaColLVTeste 2, "N", "P", "", "S", "N", "N", "E", vListviewGrupos
        End If
        If vListviewGrupos.ListItems.Count > 0 Then ajusta_LVTeste vListviewGrupos
    End If
    
'-- OS FECHAMENTO - PERMISSÃO DE COLABORADORES
    If QualLV = 15 Then
        Set vListviewPermissoes = vListViewPrincipal
        Set chamaForm = New frmPerColab
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "PFUNC", "PPESSOA", "tbautfechaos", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " PERMISSAO.CHAPA, " & vbCrLf
                SqlLV = SqlLV & " PERMISSAO.NOME, " & vbCrLf
                SqlLV = SqlLV & " PERMISSAO.ATIVO " & vbCrLf
                SqlLV = SqlLV & "FROM ( " & vbCrLf
                SqlLV = SqlLV & " SELECT TOP 500 " & vbCrLf
                SqlLV = SqlLV & "     A.CHAPA,B.NOME, " & vbCrLf
                SqlLV = SqlLV & "     CASE WHEN C.CHAPA IS NOT NULL THEN 'S' ELSE 'N' END AS ATIVO " & vbCrLf
                SqlLV = SqlLV & " FROM " & vBancoTotvs & ".DBO.PFUNC AS A " & vbCrLf
                SqlLV = SqlLV & " INNER JOIN  " & vBancoTotvs & ".DBO.PPESSOA AS B ON " & vbCrLf
                SqlLV = SqlLV & "     A.CODSITUACAO IN('A','F','P','Z') AND " & vbCrLf
                SqlLV = SqlLV & "     A.CODPESSOA = B.CODIGO AND " & vbCrLf
                SqlLV = SqlLV & "     CAST(A.CHAPA AS INT)> 10 " & vbCrLf
                SqlLV = SqlLV & " LEFT JOIN TBAUTFECHAOS AS C ON " & vbCrLf
                SqlLV = SqlLV & "     A.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI = C.CHAPA " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & " UNION " & vbCrLf
                SqlLV = SqlLV & " " & vbCrLf
                SqlLV = SqlLV & " SELECT " & vbCrLf
                SqlLV = SqlLV & "     A.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
                SqlLV = SqlLV & "     A.NOME COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
                SqlLV = SqlLV & "     CASE WHEN B.CHAPA IS NOT NULL THEN 'S' ELSE 'N' END AS ATIVO " & vbCrLf
                SqlLV = SqlLV & " FROM TBTERCEIRIZADOS AS A " & vbCrLf
                SqlLV = SqlLV & " LEFT JOIN TBAUTFECHAOS AS B ON " & vbCrLf
                SqlLV = SqlLV & "     A.CHAPA = B.CHAPA AND " & vbCrLf
                SqlLV = SqlLV & "     A.ATIVO = 'S' " & vbCrLf
                SqlLV = SqlLV & "      ) AS PERMISSAO " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " PERMISSAO.CHAPA IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & "ORDER BY PERMISSAO.CHAPA"
            End If
            
                'SqlLV = "select TOP " & LimiteLinhas & " a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from " & vBancoTotvs & ".dbo.PFUNC as a inner join " & vBancoTotvs & ".dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 " & _
                '                                  "left join tbautfechaos as c on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI = c.chapa where a.CHAPA > 0 union " & _
                '                                  "select a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI,a.nome COLLATE SQL_Latin1_General_CP1_CI_AI,case when b.chapa is not null then 'S' else 'N' end as ativo from tbTerceirizados as a left join tbautfechaos as b on a.chapa = b.chapa and a.ativo = 'S' ORDER BY a.chapa"
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Chapa", "Nome", "Permissão", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewPermissoes
        MontaDadosLVTeste "N", vListviewPermissoes
        If checaFiltro = True Then
            PersonaColLVTeste 2, "N", "P", "", "S", "N", "N", "E", vListviewPermissoes
        End If
        If vListviewPermissoes.ListItems.Count > 0 Then ajusta_LVTeste vListviewPermissoes
    End If
    
'-- LF - Relatório Liberação de Fabricação
    If QualLV = 16 Then
        Set vListviewRelInsp = vListViewPrincipal
        Set chamaForm = New frmRelInsp
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbProjetos", "tbFO", "tbCliFor", "tbFCE", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODPROJETO, " & vbCrLf
                SqlLV = SqlLV & " A.FCE, " & vbCrLf
                SqlLV = SqlLV & " A.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " C.NOME, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS = 1 THEN " & vbCrLf
                SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS IS NULL THEN " & vbCrLf
                SqlLV = SqlLV & "         'DUVIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS = 2 THEN " & vbCrLf
                SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & "     END AS STATUS " & vbCrLf
                SqlLV = SqlLV & "FROM TBPROJETOS AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFO AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.FCE=B.FCE " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBCLIFOR AS C ON " & vbCrLf
                SqlLV = SqlLV & " B.CODCLIFOR = C.CODCLIFOR " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS D ON " & vbCrLf
                SqlLV = SqlLV & " B.FCE = D.FCE " & vbCrLf
                SqlLV = SqlLV & "WHERE A.FCE > 2000 " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.FCE DESC,A.DESCRICAO"
                'SqlLV = "select top " & LimiteLinhas & " a.codprojeto,a.fce,a.projeto,c.nome,CASE WHEN d.status = 0 THEN 'ANDAMENTO' WHEN d.status = 1 THEN 'CONCLUIDA' WHEN d.status IS NULL THEN 'DUVIDA' WHEN d.status = 2 THEN 'PARALIZADA' END AS STATUS from tbProjetos as a inner join tbFO as b on a.fce=b.fce inner join tbclifor as c on b.codclifor = c.codclifor inner join tbFCE as d on b.fce = d.fce where a.fce > 2000 Order by a.fce desc,a.descricao"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "Nome Cliente", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewRelInsp
        MontaDadosLVTeste "S", vListviewRelInsp
        If checaFiltro = True Then
            PersonaColLVTeste 4, "N", "P", "", "S", "N", "N", "E", vListviewRelInsp
        End If
        If vListviewRelInsp.ListItems.Count > 0 Then ajusta_LVTeste vListviewRelInsp
    End If
    
'-- RO - Relatório de Expedição
    If QualLV = 17 Then
        Set vListviewRelExpedicao = vListViewPrincipal
        Set chamaForm = frmRelExp
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbProjetos", "tbFO", "tbCliFor", "tbFCE", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODPROJETO, " & vbCrLf
                SqlLV = SqlLV & " A.FCE, " & vbCrLf
                SqlLV = SqlLV & " A.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " C.NOME, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS = 0 THEN 'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS = 1 THEN 'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS IS NULL THEN 'DUVIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN D.STATUS = 2 THEN 'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS " & vbCrLf
                SqlLV = SqlLV & "FROM TBPROJETOS AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFO AS B ON " & vbCrLf
                SqlLV = SqlLV & " A.FCE=B.FCE " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBCLIFOR AS C ON " & vbCrLf
                SqlLV = SqlLV & " B.CODCLIFOR = C.CODCLIFOR " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBFCE AS D ON " & vbCrLf
                SqlLV = SqlLV & " B.FCE = D.FCE " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " A.FCE > 2000 " & vbCrLf
                SqlLV = SqlLV & "ORDER BY " & vbCrLf
                SqlLV = SqlLV & " A.FCE DESC, " & vbCrLf
                SqlLV = SqlLV & " A.DESCRICAO"
                'SqlLV = "select top " & LimiteLinhas & " a.codprojeto,a.fce,a.projeto,c.nome,CASE WHEN d.status = 0 THEN 'ANDAMENTO' WHEN d.status = 1 THEN 'CONCLUIDA' WHEN d.status IS NULL THEN 'DUVIDA' WHEN d.status = 2 THEN 'PARALIZADA' END AS STATUS from tbProjetos as a inner join tbFO as b on a.fce=b.fce inner join tbclifor as c on b.codclifor = c.codclifor inner join tbFCE as d on b.fce = d.fce where a.fce > 2000 Order by a.fce desc,a.descricao"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "Nome Cliente", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewRelExpedicao
        MontaDadosLVTeste "S", vListviewRelExpedicao
        If checaFiltro = True Then
            PersonaColLVTeste 4, "N", "P", "", "S", "N", "N", "E", vListviewRelExpedicao
        End If
        If vListviewRelExpedicao.ListItems.Count > 0 Then ajusta_LVTeste vListviewRelExpedicao
    End If
    
'-- IMPRESSAO DOS RELATÓRIOS DE EXPEDIÇÃO
    If QualLV = 18 Then
        Set vListviewImpExpedicao = vListViewPrincipal
        Set chamaForm = New FCRExpedicao
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            
            carregaTABS "tbRelInspExp", "rbProjetos", "tbFCE", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODREL, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.FCE = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         NULL " & vbCrLf
                SqlLV = SqlLV & "     ELSE " & vbCrLf
                SqlLV = SqlLV & "         A.FCE " & vbCrLf
                SqlLV = SqlLV & " END AS FCE, " & vbCrLf
                SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " B.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.DATAREL, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.STATUSIMP = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'NÃO IMPRESSO' " & vbCrLf
                SqlLV = SqlLV & "     ELSE " & vbCrLf
                SqlLV = SqlLV & "         'IMPRESSO' " & vbCrLf
                SqlLV = SqlLV & " END, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS = 1 THEN " & vbCrLf
                SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS IS NULL THEN " & vbCrLf
                SqlLV = SqlLV & "         'DUVIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS = 2 THEN " & vbCrLf
                SqlLV = SqlLV & "         'PARALIZADA' END " & vbCrLf
                SqlLV = SqlLV & " AS STATUS " & vbCrLf
                SqlLV = SqlLV & "FROM TBRELINSPEXP AS A " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON A.FCE = B.FCE AND A.CODPROJETO = B.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBFCE AS C ON B.FCE = C.FCE " & vbCrLf
                SqlLV = SqlLV & "WHERE A.TIPOREL = 11"
                'SqlLV = "select top " & LimiteLinhas & " a.codrel,case when a.fce = 0 then NULL else a.fce end as FCE,b.projeto,b.descricao,a.datarel,case when a.statusimp = 0 then 'Não impresso' else 'Impresso' end,CASE WHEN c.status = 0 THEN 'ANDAMENTO' WHEN c.status = 1 THEN 'CONCLUIDA' WHEN c.status IS NULL THEN 'DUVIDA' WHEN c.status = 2 THEN 'PARALIZADA' END AS STATUS from tbRelInspExp as a left join tbProjetos as b on a.fce = b.fce and a.codprojeto = b.codprojeto left join tbFCE as c on b.fce = c.fce where a.tiporel = 11"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Nº Relatório", "FCE", "Projeto", "Descrição", "Data emissão", "Status Impressão", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewImpExpedicao
        MontaDadosLVTeste "S", vListviewImpExpedicao
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "E", vListviewImpExpedicao
            PersonaColLVTeste 6, "N", "P", "", "S", "N", "N", "E", vListviewImpExpedicao
        End If
        If vListviewImpExpedicao.ListItems.Count > 0 Then ajusta_LVTeste vListviewImpExpedicao
    End If

'-- IMPRESSAO DOS RELATÓRIOS DE INSPEÇÃO (QUALIDADE)
    If QualLV = 19 Then
        Set vListviewImpInspecao = vListViewPrincipal
        Set chamaForm = New FCRLibFab
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbRelInspExp", "tbProjetos", "tbFCE", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CODREL, " & vbCrLf
                SqlLV = SqlLV & " A.FCE, " & vbCrLf
                SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
                SqlLV = SqlLV & " B.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " A.DATAREL, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.STATUSIMP = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'NÃO IMPRESSO' " & vbCrLf
                SqlLV = SqlLV & "     ELSE 'IMPRESSO' " & vbCrLf
                SqlLV = SqlLV & " END, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN A.TIPOREL = 3 THEN " & vbCrLf
                SqlLV = SqlLV & "         'FABRICAÇÃO' " & vbCrLf
                SqlLV = SqlLV & "     ELSE " & vbCrLf
                SqlLV = SqlLV & "         'PINTURA' " & vbCrLf
                SqlLV = SqlLV & " END, " & vbCrLf
                SqlLV = SqlLV & " CASE " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS = 0 THEN " & vbCrLf
                SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS = 1 THEN " & vbCrLf
                SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS IS NULL THEN " & vbCrLf
                SqlLV = SqlLV & "         'DUVIDA' " & vbCrLf
                SqlLV = SqlLV & "     WHEN C.STATUS = 2 THEN " & vbCrLf
                SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
                SqlLV = SqlLV & " END AS STATUS " & vbCrLf
                SqlLV = SqlLV & "FROM TBRELINSPEXP AS A " & vbCrLf
                SqlLV = SqlLV & "INNER JOIN TBPROJETOS AS B ON A.FCE = B.FCE AND A.CODPROJETO = B.CODPROJETO " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN TBFCE AS C ON B.FCE = C.FCE " & vbCrLf
                SqlLV = SqlLV & "WHERE A.TIPOREL < 11"
            End If
                
                'SqlLV = "select top " & LimiteLinhas & " a.codrel,a.fce,b.projeto,b.descricao, a.datarel,case when a.statusimp = 0 then 'Não impresso' else 'Impresso' end,case when a.tiporel = 3 then 'FABRICAÇÃO' else 'PINTURA' end,CASE WHEN c.status = 0 THEN 'ANDAMENTO' WHEN c.status = 1 THEN 'CONCLUIDA' WHEN c.status IS NULL THEN 'DUVIDA' WHEN c.status = 2 THEN 'PARALIZADA' END AS STATUS from tbRelInspExp as a inner join tbProjetos as b on a.fce = b.fce and a.codprojeto = b.codprojeto left join tbFCE as c on b.fce = c.fce where a.tiporel < 11"
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Nº Relatório", "FCE", "Projeto", "Descrição", "Data emissão", "Status Impressão", "Tipo", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewImpInspecao
        MontaDadosLVTeste "S", vListviewImpInspecao
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "E", vListviewImpInspecao
            PersonaColLVTeste 7, "N", "P", "", "S", "N", "N", "E", vListviewImpInspecao
        End If
        If vListviewImpInspecao.ListItems.Count > 0 Then ajusta_LVTeste vListviewImpInspecao
    End If
    
'-- FATURAMENTO POR FCE
    If QualLV = 20 Then
        Set vListviewFaturamentoFCE = vListViewPrincipal
        Set chamaForm = New FCRFatFCE
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1 'Com quantas colunas que a varglobal irá trabalhar
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "TTB3", "TMOV", "TBFCE", "FLAN", "TBPEDIDOS", "", "", "", "", ""
'            If FiltroGeral = "Em aberto" Then SqlLV = "SELECT TOP " & LimiteLinhas & " T1.DESCRICAO,T1.CODTB3FAT,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO-isnull(T4.DEVOLVIDO,0) VALOR_BRUTO,T1.VALOR_LIQUIDO-isnull(T4.DEVOLVIDO,0) VALOR_LIQUIDO,T1.DTCRIACAO,(T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL,T2.VALOR_BAIXADO as VALOR_BAIXADO, " & _
'                                                  "((T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-isnull(T5.ADIANTADO,0)-isnull(T6.CANCELADO,0) AS VALOR_RECEBER,T3.PESO,T3.VALOR_TOTAL,CASE WHEN T1.status = 0 THEN 'ANDAMENTO' WHEN T1.status = 1 THEN 'CONCLUIDA' WHEN T1.status IS NULL THEN 'DUVIDA' WHEN T1.status = 2 THEN 'PARALIZADA' END AS STATUS FROM (select MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO, " & _
'                                                  "SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO,MAX(c.status) as status FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' left join tbFCE as C on a.CODTB3FAT = c.fce GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) " & _
'                                                  ") T1 LEFT JOIN (SELECT B.CODTB3FAT,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY B.CODTB3FAT) T2 " & _
'                                                  "ON T1.CODTB3FAT = T2.CODTB3FAT LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from " & sDatabaseName & ".DBO.tbPedidos AS A group by a.fce) T3 ON T1.CODTB3FAT = T3.FCE LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as DEVOLVIDO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT " & _
'                                                  "where B.CODTMV in ('1.2.15','1.2.17') and B.STATUS = 'F' group by B.CODTB3FAT) T4 ON T1.CODTB3FAT = T4.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as ADIANTADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.25') group by B.CODTB3FAT) T5 ON T1.CODTB3FAT = T5.CODTB3FAT " & _
'                                                  "LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as CANCELADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT where B.STATUS = 'C' and B.CODTMV in ('2.2.01','2.2.05','1.2.15','1.2.17') group by b.CODTB3FAT) T6 ON T1.CODTB3FAT = T6.CODTB3FAT where T2.VALOR_RECEBER > 0 or T2.VALOR_RECEBER is null ORDER BY T1.CODTB3FAT"
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " T1.DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT, " & vbCrLf
                SqlLV = SqlLV & " T1.PESO_LIQUIDO, " & vbCrLf
                SqlLV = SqlLV & " T1.PESO_BRUTO, " & vbCrLf
                SqlLV = SqlLV & " T1.VALOR_BRUTO-ISNULL(T4.DEVOLVIDO,0) VALOR_BRUTO, " & vbCrLf
                SqlLV = SqlLV & " T1.VALOR_LIQUIDO-ISNULL(T4.DEVOLVIDO,0) VALOR_LIQUIDO, " & vbCrLf
                SqlLV = SqlLV & " T1.DTCRIACAO, " & vbCrLf
                SqlLV = SqlLV & " (T2.VALOR_ORIGINAL+ISNULL(T4.DEVOLVIDO,0)+ISNULL(T5.ADIANTADO,0)+ISNULL(T6.CANCELADO,0))-(ISNULL(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL, " & vbCrLf
                SqlLV = SqlLV & " T2.VALOR_BAIXADO AS VALOR_BAIXADO, " & vbCrLf
                SqlLV = SqlLV & " ((T2.VALOR_ORIGINAL+ISNULL(T4.DEVOLVIDO,0)+ISNULL(T5.ADIANTADO,0)+ISNULL(T6.CANCELADO,0))-(ISNULL(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-ISNULL(T5.ADIANTADO,0)-ISNULL(T6.CANCELADO,0) AS VALOR_RECEBER, " & vbCrLf
                SqlLV = SqlLV & " T3.PESO, " & vbCrLf
                SqlLV = SqlLV & " T3.VALOR_TOTAL, " & vbCrLf
                SqlLV = SqlLV & " CASE WHEN T1.STATUS = 0 THEN 'ANDAMENTO' WHEN T1.STATUS = 1 THEN 'CONCLUIDA' WHEN T1.STATUS IS NULL THEN 'DUVIDA' WHEN T1.STATUS = 2 THEN 'PARALIZADA' END AS STATUS " & vbCrLf
                SqlLV = SqlLV & "FROM " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & "     SELECT " & vbCrLf
                SqlLV = SqlLV & "         MAX(B.IDMOV) AS IDMOV, " & vbCrLf
                SqlLV = SqlLV & "         A.CODTB3FAT, " & vbCrLf
                SqlLV = SqlLV & "         SUBSTRING(A.DESCRICAO,1,50) AS DESCRICAO, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.PESOBRUTO) AS PESO_BRUTO, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.VALORBRUTO) AS VALOR_LIQUIDO, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.VALORLIQUIDO) AS VALOR_BRUTO, " & vbCrLf
                SqlLV = SqlLV & "         MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) AS DTCRIACAO, " & vbCrLf
                SqlLV = SqlLV & "         MAX(C.STATUS) AS STATUS " & vbCrLf
                SqlLV = SqlLV & "     FROM " & vBancoTotvs & ".DBO.TTB3 AS A " & vbCrLf
                SqlLV = SqlLV & "     LEFT JOIN " & vBancoTotvs & ".DBO.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV IN ('2.2.01','2.2.05') AND B.STATUS <> 'C' " & vbCrLf
                SqlLV = SqlLV & "     LEFT JOIN TBFCE AS C ON A.CODTB3FAT = C.FCE " & vbCrLf
                SqlLV = SqlLV & "     GROUP BY A.CODTB3FAT,SUBSTRING(A.DESCRICAO,1,50) " & vbCrLf
                SqlLV = SqlLV & " ) AS T1 " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & "     SELECT " & vbCrLf
                SqlLV = SqlLV & "         B.CODTB3FAT, " & vbCrLf
                SqlLV = SqlLV & "         SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL, " & vbCrLf
                SqlLV = SqlLV & "         SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO, " & vbCrLf
                SqlLV = SqlLV & "         SUM(C.VALORORIGINAL-C.VALORBAIXADO) AS VALOR_RECEBER " & vbCrLf
                SqlLV = SqlLV & "     FROM " & vBancoTotvs & ".DBO.TMOV AS B " & vbCrLf
                SqlLV = SqlLV & "     INNER JOIN " & vBancoTotvs & ".DBO.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV IN ('2.2.01','2.2.05') AND B.STATUS <> 'C' " & vbCrLf
                SqlLV = SqlLV & "     GROUP BY B.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & "     ) AS T2 ON " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT = T2.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & "     SELECT " & vbCrLf
                SqlLV = SqlLV & "         A.FCE AS FCE, " & vbCrLf
                SqlLV = SqlLV & "         SUM(A.PESO) AS PESO, " & vbCrLf
                SqlLV = SqlLV & "         SUM(A.TOTAL) AS VALOR_TOTAL " & vbCrLf
                SqlLV = SqlLV & "     FROM " & sDatabaseName & ".DBO.TBPEDIDOS AS A " & vbCrLf
                SqlLV = SqlLV & "     GROUP BY A.FCE " & vbCrLf
                SqlLV = SqlLV & " ) T3 ON " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT = T3.FCE " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & "     SELECT B.CODTB3FAT, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.VALORLIQUIDO) AS DEVOLVIDO " & vbCrLf
                SqlLV = SqlLV & "     FROM " & vBancoTotvs & ".DBO.TTB3 AS A " & vbCrLf
                SqlLV = SqlLV & "     LEFT JOIN " & vBancoTotvs & ".DBO.TMOV AS B ON " & vbCrLf
                SqlLV = SqlLV & "         A.CODTB3FAT = B.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & "     WHERE " & vbCrLf
                SqlLV = SqlLV & "         B.CODTMV IN ('1.2.15','1.2.17') AND " & vbCrLf
                SqlLV = SqlLV & "         B.STATUS = 'F' " & vbCrLf
                SqlLV = SqlLV & "     GROUP BY B.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & " ) T4 ON " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT = T4.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & "     SELECT " & vbCrLf
                SqlLV = SqlLV & "         B.CODTB3FAT, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.VALORLIQUIDO) AS ADIANTADO " & vbCrLf
                SqlLV = SqlLV & "     FROM " & vBancoTotvs & ".DBO.TTB3 AS A " & vbCrLf
                SqlLV = SqlLV & "     LEFT JOIN " & vBancoTotvs & ".DBO.TMOV AS B ON " & vbCrLf
                SqlLV = SqlLV & "         A.CODTB3FAT = B.CODTB3FAT AND " & vbCrLf
                SqlLV = SqlLV & "         B.CODTMV IN ('2.2.25') " & vbCrLf
                SqlLV = SqlLV & "     GROUP BY B.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & " ) T5 ON " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT = T5.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
                SqlLV = SqlLV & " ( " & vbCrLf
                SqlLV = SqlLV & "     SELECT " & vbCrLf
                SqlLV = SqlLV & "         B.CODTB3FAT, " & vbCrLf
                SqlLV = SqlLV & "         SUM(B.VALORLIQUIDO) AS CANCELADO " & vbCrLf
                SqlLV = SqlLV & "     FROM " & vBancoTotvs & ".DBO.TTB3 AS A " & vbCrLf
                SqlLV = SqlLV & "     LEFT JOIN " & vBancoTotvs & ".DBO.TMOV AS B ON " & vbCrLf
                SqlLV = SqlLV & "         A.CODTB3FAT = B.CODTB3FAT AND" & vbCrLf
                SqlLV = SqlLV & "         B.STATUS = 'C' AND " & vbCrLf
                SqlLV = SqlLV & "         B.CODTMV IN ('2.2.01','2.2.05','1.2.15','1.2.17') " & vbCrLf
                SqlLV = SqlLV & "     GROUP BY B.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & " ) T6 ON " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT = T6.CODTB3FAT " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " T1.IDMOV IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & "ORDER BY " & vbCrLf
                SqlLV = SqlLV & " T1.CODTB3FAT DESC"
                
'                SqlLV = "SELECT TOP " & LimiteLinhas & " T1.DESCRICAO,T1.CODTB3FAT,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO-isnull(T4.DEVOLVIDO,0) VALOR_BRUTO,T1.VALOR_LIQUIDO-isnull(T4.DEVOLVIDO,0) VALOR_LIQUIDO,T1.DTCRIACAO,(T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL,T2.VALOR_BAIXADO as VALOR_BAIXADO, " & _
'                                                  "((T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-isnull(T5.ADIANTADO,0)-isnull(T6.CANCELADO,0) AS VALOR_RECEBER,T3.PESO,T3.VALOR_TOTAL,CASE WHEN T1.status = 0 THEN 'ANDAMENTO' WHEN T1.status = 1 THEN 'CONCLUIDA' WHEN T1.status IS NULL THEN 'DUVIDA' WHEN T1.status = 2 THEN 'PARALIZADA' END AS STATUS FROM (select MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO, " & _
'                                                  "SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO,MAX(c.status) as status FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' left join tbFCE as C on a.CODTB3FAT = c.fce GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) " & _
'                                                  ") T1 LEFT JOIN (SELECT B.CODTB3FAT,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY B.CODTB3FAT) T2 " & _
'                                                  "ON T1.CODTB3FAT = T2.CODTB3FAT LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from " & sDatabaseName & ".DBO.tbPedidos AS A group by a.fce) T3 ON T1.CODTB3FAT = T3.FCE LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as DEVOLVIDO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT " & _
'                                                  "where B.CODTMV in ('1.2.15','1.2.17') and B.STATUS = 'F' group by B.CODTB3FAT) T4 ON T1.CODTB3FAT = T4.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as ADIANTADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.25') group by B.CODTB3FAT) T5 ON T1.CODTB3FAT = T5.CODTB3FAT " & _
'                                                  "LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as CANCELADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT where B.STATUS = 'C' and B.CODTMV in ('2.2.01','2.2.05','1.2.15','1.2.17') group by b.CODTB3FAT) T6 ON T1.CODTB3FAT = T6.CODTB3FAT ORDER BY T1.CODTB3FAT"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Descrição", "FCE", "Peso Líquido (FAT)", "Peso Bruto (FAT)", "Valor Bruto (FAT)", "Valor Líquido (FAT)", "Data Cadastro(FCE)", "Valor Original (FIN)", "Valor Baixado (FIN)", "Valor Receber (FIN)", "Peso (COM)", "Valor Vendido (COM)", "Status", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewFaturamentoFCE
        MontaDadosLVTeste "N", vListviewFaturamentoFCE
        If checaFiltro = True Then
            PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "E", vListviewFaturamentoFCE
            PersonaColLVTeste 2, "N", "N", "", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 3, "N", "N", "", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 4, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 5, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 7, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 8, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 9, "S", "S", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 10, "N", "N", "", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 11, "S", "S", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
            PersonaColLVTeste 12, "N", "P", "", "S", "N", "N", "E", vListviewFaturamentoFCE
        End If
        If vListviewFaturamentoFCE.ListItems.Count > 0 Then ajusta_LVTeste vListviewFaturamentoFCE
    End If
    
'-- TERCEIRIZADOS
    If QualLV = 21 Then
        Set vListviewTerceiros = vListViewPrincipal
        Set chamaForm = New frmTerceirizados
        LegendaExc = Formulario 'Usado na mensagem de exclusão
        indiceVarGlobal = 1
        If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
            MontaFiltro
            If FiltroGeral = "" Then
                frmFiltro.Show 1
            Else
                filtroPadrao
            End If
            carregaTABS "tbusuarios", "tbgrupo", "", "", "", "", "", "", "", ""
            If FiltroGeral = "Todos" Then
                SqlLV = ""
                SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
                SqlLV = SqlLV & " A.CHAPA, " & vbCrLf
                SqlLV = SqlLV & " A.NOME, " & vbCrLf
                SqlLV = SqlLV & " A.IDSETOR, " & vbCrLf
                SqlLV = SqlLV & " A.SETOR, " & vbCrLf
                SqlLV = SqlLV & " A.IDFUNCAO, " & vbCrLf
                SqlLV = SqlLV & " A.FUNCAO, " & vbCrLf
                SqlLV = SqlLV & " A.IDCC, " & vbCrLf
                SqlLV = SqlLV & " A.NMCC, " & vbCrLf
                SqlLV = SqlLV & " A.EMPRESA, " & vbCrLf
                SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
                SqlLV = SqlLV & " A.DATACONTRATOINI, " & vbCrLf
                SqlLV = SqlLV & " A.DATACONTRATOFIM, " & vbCrLf
                SqlLV = SqlLV & " A.ATIVO " & vbCrLf
                SqlLV = SqlLV & "FROM TBTERCEIRIZADOS AS A " & vbCrLf
                SqlLV = SqlLV & "WHERE " & vbCrLf
                SqlLV = SqlLV & " A.CHAPA IS NOT NULL " & vbCrLf
                SqlLV = SqlLV & "ORDER BY A.CHAPA DESC"
'                SqlLV = "select TOP " & LimiteLinhas & " a.chapa,a.nome,a.idsetor,a.setor,a.idfuncao,a.funcao,a.idcc,a.nmcc,a.empresa,a.datacadastro,a.datacontratoini,a.datacontratofim,a.ativo from tbTerceirizados as a"
            End If
        Else
            If MeuLV.Visible = True Then Unload MeuLV
        End If
        MontaCabLV "Código", "Nome do usuário", "ID Setor", "Nome Setor", "ID Função", "Nome Função", "ID CC", "Nome CC", "Empresa", "D. Cadastro", "D. Contrato ini.", "D. Contrato Fim", "Ativo", "", "", "", "", "", "", "", ""
        MontaCabecalhoLVTeste vListviewTerceiros
        MontaDadosLVTeste "S", vListviewTerceiros
        If checaFiltro = True Then
            PersonaColLVTeste 12, "N", "P", "", "S", "N", "N", "E", vListviewTerceiros
        End If
        If vListviewTerceiros.ListItems.Count > 0 Then ajusta_LVTeste vListviewTerceiros
    End If
    vListViewPrincipal.Visible = True
    'construirBotaoClose frmPesqGeralTeste2.SSTab1
    Exit Function
TrataErro:
    montaLV1 = False
    Resume Next

End Function

'Public Sub MontaLV(QualLV As Integer)
'        'On Error GoTo TrataErro
'        vTimer = False
'
'        'TIPO DE MATERIAIS
'        If QualLV = 0 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmTipoMat
'            Formulario = "Tipo de materiais"
'            LegendaExc = "Tipo de materiais" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 2
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbTipoMat", "", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "Select a.codigo,a.descricao,a.ativo from tbTipoMat as a "
'                If FiltroGeral = "Ativos" Then SqlLV = "Select a.codigo,a.descricao,a.ativo from tbTipoMat as a where a.ativo = 'S'"
'                'If FiltroGeral = "Não ativos" Then SqlLV = "Select a.codigo,a.descricao,a.ativo from tbTipoMat as a where a.ativo <> 'S'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(9).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Código", "Descrição", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Tipo de materiais"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'
'            If checaFiltro = True Then
'                'PersonaColLV 0, "N", "N", "", "N", "N", "N", "D"
'                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            'Set MeuLV.cmdconsulta(6).PictureNormal = MeuLV.ImageList1.ListImages(4).Picture
'            MeuLV.Visible = True
'            Exit Sub
'        'CLIENTES
'        ElseIf QualLV = 1 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'
'            Set chamaForm = New frmClientes
'            Formulario = "Clientes"
'            LegendaExc = "Clientes" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbclifor", "", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "Select a.codclifor,a.nome,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo from tbclifor as a Order by a.codclifor "
'                If FiltroGeral = "Ativos" Then SqlLV = "Select a.codclifor,a.nome,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo  from tbclifor as a where a.ativo='S' Order by a.codclifor "
'                If FiltroGeral = "Não ativos" Then SqlLV = "Select a.codclifor,a.nome,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo  from tbclifor as a where a.ativo<>'S' Order by a.codclifor "
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(9).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Código", "Nome", "Endereço", "CEP", "Bairro", "Cidade", "UF", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Clientes"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
''                PersonaColLV 4, "S", "S", "%", "N", "N", "S", "D"
'                PersonaColLV 7, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'MOVIMENTAÇÕES - OS
'        ElseIf QualLV = 2 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmAtividades
'            Formulario = "Paradas - OS"
'            LegendaExc = "Movimentação" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbparadas", "", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "Select a.idparada,a.tipo,a.codigo,a.nmparada,a.descricao,a.ativo from tbparadas as a"
'                If FiltroGeral = "Ativos" Then SqlLV = "Select a.idparada,a.tipo,a.codigo,a.nmparada,a.descricao,a.ativo from tbparadas as a where a.ativo = 'S'"
'                If FiltroGeral = "Não ativos" Then SqlLV = "Select a.idparada,a.tipo,a.codigo,a.nmparada,a.descricao,a.ativo from tbparadas as a where a.ativo <> 'S'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "ID", "Tipo", "Código", "Nome", "Descrição", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Paradas - OS"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                'PersonaColLV 0, "N", "N", "", "N", "N", "N", "D"
'                PersonaColLV 5, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
''        'TRANSPORTADORA
'        ElseIf QualLV = 3 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmTransportes
'            Formulario = "Transportadoras"
'            LegendaExc = "Transportadoras" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
''                If FiltroGeral = "Todos" Then SqlLV = "select a.codtransp,a.nome,a.cnpj,a.ie,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo from tbTransportadoras as a"
''                If FiltroGeral = "Ativos" Then SqlLV = "Select a.codtransp,a.nome,a.cnpj,a.ie,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo from tbTransportadoras as a where a.ativo = 'S'"
''                If FiltroGeral = "Não ativos" Then SqlLV = "Select a.codtransp,a.nome,a.cnpj,a.ie,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.ativo from tbTransportadoras as a where a.ativo <> 'S'"
'
'                carregaTABS "ttra", "", "", "", "", "", "", "", "", ""
'
'                'If FiltroGeral = "Todos" Then SqlLV = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD,a.INATIVO from CORPORERM.dbo.ttra as a"
'                If FiltroGeral = "Todos" Then SqlLV = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD,a.INATIVO from " & vBancoTotvs & ".dbo.ttra as a where a.inativo = 0 or a.inativo is null"
'                'If FiltroGeral = "Não ativos" Then SqlLV = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD,a.INATIVO from CORPORERM.dbo.ttra as a where a.inativo > 0"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Código", "Nome", "CNPJ", "IE", "Endereço", "CEP", "Bairro", "Cidade", "UF", "Ativo", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Transportadoras"
'            MontaCabecalhoLV
'            MontaDadosLV "N"
'            If checaFiltro = True Then
'                PersonaColLV 9, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'FÓRMULAS PRODUTOS
'        ElseIf QualLV = 4 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmMaterial
'            Formulario = "Fórmulas de Produtos"
'            LegendaExc = "Fórmulas de Produtos" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "TPRD", "tbMateriais", "ttb2", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select a.idprd,a.CODIGOPRD,a.NOMEFANTASIA,a.codtb2fat,c.descricao,b.formula,b.forpint from " & vBancoTotvs & ".dbo.TPRD as a left join " & sDatabaseName & ".dbo.tbMateriais as b on a.IDPRD = b.IDPRD left join " & vBancoTotvs & ".dbo.ttb2 as c on a.CODTB2FAT = c.CODTB2FAT and c.CODCOLIGADA = " & vCodcoligada & " where a.CODIGOPRD like '%%' and a.CODCOLIGADA = " & vCodcoligada
''                If FiltroGeral = "Ativos" Then SqlLV = "select a.codmaterial,a.descricao,a.unidade,a.formula,a.forpint,a.constpint,a.ativo from tbMateriais as a where a.ativo = 'S'"
''                If FiltroGeral = "Não ativos" Then SqlLV = "select a.codmaterial,a.descricao,a.unidade,a.formula,a.forpint,a.constpint,a.ativo from tbMateriais as a where a.ativo <> 'S'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "ID", "Código", "Descrição", "Cod Tipo", "Tipo Material", "Fórmula PESO", "Fórmula PINTURA", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Fórmulas de Produtos"
'            MontaCabecalhoLV
'            MontaDadosLV "S" 'Coloca zeros a esquerda na primeira coluna
'            If checaFiltro = True Then
'                'PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.cmdconsulta(6).ToolTipText = "Limpar dados da Fórmula"
'            MeuLV.Visible = True
'            Exit Sub
'        'ORÇAMENTOS
'        ElseIf QualLV = 5 Then
'
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'
'            Set chamaForm = New frmFO
'            Formulario = "Orçamentos"
'            LegendaExc = "Orçamentos" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                'If MeuLV.Visible = True Then Unload MeuLV
'                carregaTABS "tbclifor", "tbfo", "tbcontatos", "tbfce", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then 'SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo,CASE WHEN D.status = 0 THEN 'ANDAMENTO' WHEN D.status = 1 THEN 'CONCLUIDA' WHEN D.status = 2 THEN 'PARALIZADA' END AS STATUS from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato left join tbFCE as d on b.fce = d.fce"
'                        SqlLV = ""
'                        SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                        SqlLV = SqlLV & " FO.CODFO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.NOME, " & vbCrLf
'                        SqlLV = SqlLV & " FO.PEDIDO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.NOME_CONTATO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.TELEFONE, " & vbCrLf
'                        SqlLV = SqlLV & " FO.DESCRICAO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.DATAFO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.DATADEVCP, " & vbCrLf
'                        SqlLV = SqlLV & " FO.PROPOSTA, " & vbCrLf
'                        SqlLV = SqlLV & " FO.QUANTIDADE, " & vbCrLf
'                        SqlLV = SqlLV & " FO.VALORUNIT, " & vbCrLf
'                        SqlLV = SqlLV & " FO.VALORTOTAL, " & vbCrLf
'                        SqlLV = SqlLV & " FO.PEDIDO1, " & vbCrLf
'                        SqlLV = SqlLV & " CASE WHEN FO.FCE IS NULL THEN 0 ELSE FO.FCE END AS FCE, " & vbCrLf
'                        SqlLV = SqlLV & " FO.STATUSFO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.ATIVO, " & vbCrLf
'                        SqlLV = SqlLV & " FO.STATUS, " & vbCrLf
'                        SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS [TIPO FCE] " & vbCrLf
'                        SqlLV = SqlLV & "FROM ( " & vbCrLf
'                        SqlLV = SqlLV & " SELECT " & vbCrLf
'                        SqlLV = SqlLV & "     B.CODFO AS CODFO, " & vbCrLf
'                        SqlLV = SqlLV & "     A.NOME AS NOME, " & vbCrLf
'                        SqlLV = SqlLV & "     B.PEDIDO AS PEDIDO, " & vbCrLf
'                        SqlLV = SqlLV & "     C.NOME AS NOME_CONTATO, " & vbCrLf
'                        SqlLV = SqlLV & "     C.TELEFONE AS TELEFONE, " & vbCrLf
'                        SqlLV = SqlLV & "     B.DESCRICAO AS DESCRICAO, " & vbCrLf
'                        SqlLV = SqlLV & "     B.DATAFO AS DATAFO, " & vbCrLf
'                        SqlLV = SqlLV & "     B.DATADEVCP AS DATADEVCP, " & vbCrLf
'                        SqlLV = SqlLV & "     B.PROPOSTA AS PROPOSTA, " & vbCrLf
'                        SqlLV = SqlLV & "     B.QUANTIDADE AS QUANTIDADE, " & vbCrLf
'                        SqlLV = SqlLV & "     B.VALORUNIT AS VALORUNIT, " & vbCrLf
'                        SqlLV = SqlLV & "     (B.QUANTIDADE*B.VALORUNIT) AS VALORTOTAL, " & vbCrLf
'                        SqlLV = SqlLV & "     B.PEDIDO AS PEDIDO1, " & vbCrLf
'                        SqlLV = SqlLV & "     B.FCE AS FCE, " & vbCrLf
'                        SqlLV = SqlLV & "     B.STATUSFO AS STATUSFO, " & vbCrLf
'                        SqlLV = SqlLV & "     B.ATIVO ATIVO, " & vbCrLf
'                        SqlLV = SqlLV & "     CASE WHEN D.STATUS = 0 THEN 'ANDAMENTO' WHEN D.STATUS = 1 THEN 'CONCLUIDA' WHEN D.STATUS = 2 THEN 'PARALIZADA' END AS STATUS " & vbCrLf
'                        SqlLV = SqlLV & " FROM " & vbCrLf
'                        SqlLV = SqlLV & " TBCLIFOR AS A " & vbCrLf
'                        SqlLV = SqlLV & " INNER JOIN TBFO  AS B ON A.CODCLIFOR=B.CODCLIFOR " & vbCrLf
'                        SqlLV = SqlLV & " LEFT JOIN TBCONTATOS AS C ON B.CODCLIFOR = C.CODCLIFOR AND B.CODCONTATO = C.CODCONTATO " & vbCrLf
'                        SqlLV = SqlLV & " LEFT JOIN TBFCE AS D ON B.FCE = D.FCE " & vbCrLf
'                        SqlLV = SqlLV & ") AS FO " & vbCrLf
'                        SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                        SqlLV = SqlLV & " ( " & vbCrLf
'                        SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                        SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                        SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                        SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                        SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                        SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                        SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                        SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                        SqlLV = SqlLV & " " & vbCrLf
'                        SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                        SqlLV = SqlLV & " ) AS FILTRO ON " & vbCrLf
'                        SqlLV = SqlLV & "FO.FCE = FILTRO.FCE ORDER BY FO.CODFO DESC"
'
'
'                End If
'                If FiltroGeral = "Ativos" Then SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo,CASE WHEN D.status = 0 THEN 'ANDAMENTO' WHEN D.status = 1 THEN 'CONCLUIDA' WHEN D.status = 2 THEN 'PARALIZADA' END AS STATUS from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato left join tbFCE as d on b.fce = d.fce where b.ativo = 'S'"
'                If FiltroGeral = "Não ativos" Then SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo,CASE WHEN D.status = 0 THEN 'ANDAMENTO' WHEN D.status = 1 THEN 'CONCLUIDA' WHEN D.status = 2 THEN 'PARALIZADA' END AS STATUS from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato left join tbFCE as d on b.fce = d.fce where b.ativo <> 'S'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(9).Visible = True
'            MeuLV.cmdconsulta(11).Visible = True
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "FO", "Empresa", "Coleta nº", "Contato", "Fone", "Descrição", "Data Abertura", "Dev. CP", "Proposta nº", "Quant.", "Valor Unit", "Valor Total", "Pedido nº", "FCE nº", "Status FO", "Ativo", "Status FCE", "Tipo FCE", "", "", ""
'            DimensionaLV "Orçamentos"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                PersonaColLV 14, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 13, "S", "S", "", "N", "N", "N", "D"
'                PersonaColLV 15, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 16, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 17, "N", "P", "", "N", "N", "N", "E" 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Set MeuLV.cmdconsulta(9).PictureNormal = MeuLV.ImageList1.ListImages(7).Picture
'            MeuLV.cmdconsulta(9).ToolTipText = "Receber FO"
'            Set MeuLV.cmdconsulta(11).PictureNormal = MeuLV.ImageList1.ListImages(22).Picture
'            MeuLV.cmdconsulta(11).ToolTipText = "Editar FCE"
'            MeuLV.Visible = True
'            Exit Sub
''        'FCE - Ficha de Controle de Encomenda
'        ElseIf QualLV = 6 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmFCECons
'            Formulario = "FCE"
'            LegendaExc = "FCE's - Fichas de Controle de Encomenda" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 3
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                'If MeuLV.Visible = True Then Unload MeuLV
'                carregaTABS "tbfo", "tbfce", "tbclifor", "tbcontatos", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then
'                    'SqlLV = "select b.dataabertura,b.fce,c.nome[cliente],d.nome[contato],d.telefone,b.dataentrega,b.pintura,b.transporte,b.materiaprima,b.fabricacao,b.reparo,a.ativo,CASE WHEN B.status = 0 THEN 'ANDAMENTO' WHEN B.status = 1 THEN 'CONCLUIDA' WHEN B.status = 2 THEN 'PARALIZADA' END AS STATUS from tbfo as a inner join tbfce as b on b.fce = a.fce left join tbclifor as c on a.codclifor=c.codclifor left join tbcontatos as d on c.codclifor = d.codclifor and d.codcontato = a.codcontato order by b.fce"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " B.DATAABERTURA, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " C.NOME[CLIENTE], " & vbCrLf
'                    SqlLV = SqlLV & " D.NOME[CONTATO], " & vbCrLf
'                    SqlLV = SqlLV & " D.TELEFONE, " & vbCrLf
'                    SqlLV = SqlLV & " B.DATAENTREGA, " & vbCrLf
'                    SqlLV = SqlLV & " B.PINTURA, " & vbCrLf
'                    SqlLV = SqlLV & " B.TRANSPORTE, " & vbCrLf
'                    SqlLV = SqlLV & " B.MATERIAPRIMA, " & vbCrLf
'                    SqlLV = SqlLV & " B.FABRICACAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.REPARO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN B.STATUS = 0 THEN 'ANDAMENTO' WHEN B.STATUS = 1 THEN 'CONCLUIDA' WHEN B.STATUS = 2 THEN 'PARALIZADA' END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBFO AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON B.FCE = A.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBCLIFOR AS C ON A.CODCLIFOR=C.CODCLIFOR " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBCONTATOS AS D ON " & vbCrLf
'                    SqlLV = SqlLV & " C.CODCLIFOR = D.CODCLIFOR AND " & vbCrLf
'                    SqlLV = SqlLV & " D.CODCONTATO = A.CODCONTATO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY B.FCE DESC"
'                End If
'                If FiltroGeral = "Ativos" Then
'                    'SqlLV = "select b.dataabertura,b.fce,c.nome[cliente],d.nome[contato],d.telefone,b.dataentrega,b.pintura,b.transporte,b.materiaprima,b.fabricacao,b.reparo,a.ativo,CASE WHEN B.status = 0 THEN 'ANDAMENTO' WHEN B.status = 1 THEN 'CONCLUIDA' WHEN B.status = 2 THEN 'PARALIZADA' END AS STATUS from tbfo as a inner join tbfce as b on b.fce = a.fce left join tbclifor as c on a.codclifor=c.codclifor left join tbcontatos as d on c.codclifor = d.codclifor and d.codcontato = a.codcontato where a.ativo = 'S' order by b.fce"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " B.DATAABERTURA, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " C.NOME[CLIENTE], " & vbCrLf
'                    SqlLV = SqlLV & " D.NOME[CONTATO], " & vbCrLf
'                    SqlLV = SqlLV & " D.TELEFONE, " & vbCrLf
'                    SqlLV = SqlLV & " B.DATAENTREGA, " & vbCrLf
'                    SqlLV = SqlLV & " B.PINTURA, " & vbCrLf
'                    SqlLV = SqlLV & " B.TRANSPORTE, " & vbCrLf
'                    SqlLV = SqlLV & " B.MATERIAPRIMA, " & vbCrLf
'                    SqlLV = SqlLV & " B.FABRICACAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.REPARO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN B.STATUS = 0 THEN 'ANDAMENTO' WHEN B.STATUS = 1 THEN 'CONCLUIDA' WHEN B.STATUS = 2 THEN 'PARALIZADA' END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBFO AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON B.FCE = A.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBCLIFOR AS C ON A.CODCLIFOR=C.CODCLIFOR " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBCONTATOS AS D ON " & vbCrLf
'                    SqlLV = SqlLV & " C.CODCLIFOR = D.CODCLIFOR AND " & vbCrLf
'                    SqlLV = SqlLV & " D.CODCONTATO = A.CODCONTATO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE  WHERE A.ATIVO = 'S' " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY B.FCE DESC"
'                End If
'                If FiltroGeral = "Não ativos" Then
'                    'SqlLV = "select b.dataabertura,b.fce,c.nome[cliente],d.nome[contato],d.telefone,b.dataentrega,b.pintura,b.transporte,b.materiaprima,b.fabricacao,b.reparo,a.ativo,CASE WHEN B.status = 0 THEN 'ANDAMENTO' WHEN B.status = 1 THEN 'CONCLUIDA' WHEN B.status = 2 THEN 'PARALIZADA' END AS STATUS from tbfo as a inner join tbfce as b on b.fce = a.fce left join tbclifor as c on a.codclifor=c.codclifor left join tbcontatos as d on c.codclifor = d.codclifor and d.codcontato = a.codcontato where a.ativo <> 'S' order by b.fce"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " B.DATAABERTURA, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " C.NOME[CLIENTE], " & vbCrLf
'                    SqlLV = SqlLV & " D.NOME[CONTATO], " & vbCrLf
'                    SqlLV = SqlLV & " D.TELEFONE, " & vbCrLf
'                    SqlLV = SqlLV & " B.DATAENTREGA, " & vbCrLf
'                    SqlLV = SqlLV & " B.PINTURA, " & vbCrLf
'                    SqlLV = SqlLV & " B.TRANSPORTE, " & vbCrLf
'                    SqlLV = SqlLV & " B.MATERIAPRIMA, " & vbCrLf
'                    SqlLV = SqlLV & " B.FABRICACAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.REPARO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN B.STATUS = 0 THEN 'ANDAMENTO' WHEN B.STATUS = 1 THEN 'CONCLUIDA' WHEN B.STATUS = 2 THEN 'PARALIZADA' END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBFO AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON B.FCE = A.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBCLIFOR AS C ON A.CODCLIFOR=C.CODCLIFOR " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBCONTATOS AS D ON " & vbCrLf
'                    SqlLV = SqlLV & " C.CODCLIFOR = D.CODCLIFOR AND " & vbCrLf
'                    SqlLV = SqlLV & " D.CODCONTATO = A.CODCONTATO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE  WHERE A.ATIVO <> 'S' " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY B.FCE DESC"
'                End If
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Data abertura", "FCE", "Cliente", "Contato", "Fone", "Data entrega", "Pintura", "Transporte", "Matéria-prima", "Fabricação", "Reparo", "Ativo", "Status FCE", "Tipo FCE", "", "", "", "", "", "", ""
'            DimensionaLV "FCE's - Fichas de Controle de Encomenda"
'            MontaCabecalhoLV
'            MontaDadosLV "N"
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "S", "", "N", "N", "N", "D"
'                PersonaColLV 11, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 12, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 13, "N", "P", "", "N", "N", "N", "E" 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(8).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Nova LM - Lista de Materiais"
'            MeuLV.cmdconsulta(5).ToolTipText = "Consultar FCE - Ficha de Controle de Encomenda"
'            MeuLV.Visible = True
'            Exit Sub
'        'CADASTRO DE DESENHOS
'        ElseIf QualLV = 7 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmDesenhos
'            Formulario = "Cadastro de Desenhos"
'            LegendaExc = "Cadastro de Desenhos" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                'If FiltroGeral = "" Then frmFiltro.Show 1
'                'If MeuLV.Visible = True Then Unload MeuLV
'                carregaTABS "tbDesenhos", "tbProjetos", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then
'                    'SqlLV = "Select top " & LimiteLinhas & " a.iddesenho,a.desenho,a.revisao,b.fce,b.projeto,a.datacadastro,a.tipo,a.ativo from  tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto where a.codcoligada = '" & vCodcoligada & "' order by b.fce desc,b.projeto"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " A.IDDESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.REVISAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
'                    SqlLV = SqlLV & " A.TIPO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBDESENHOS AS A " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON " & vbCrLf
'                    SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
'                    SqlLV = SqlLV & " " & vbCrLf
'                    SqlLV = SqlLV & "WHERE A.CODCOLIGADA = '" & vCodcoligada & "' " & vbCrLf
'                    SqlLV = SqlLV & " " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE DESC,B.PROJETO"
'                End If
'                If FiltroGeral = "Ativos" Then
'                    'SqlLV = "Select top " & LimiteLinhas & " a.iddesenho,a.desenho,a.revisao,b.fce,b.projeto,a.datacadastro,a.tipo,a.ativo from  tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto where a.codcoligada = '" & vCodcoligada & "' and a.ativo='S' order by b.fce desc,b.projeto"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " A.IDDESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.REVISAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
'                    SqlLV = SqlLV & " A.TIPO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBDESENHOS AS A " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON " & vbCrLf
'                    SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
'                    SqlLV = SqlLV & " " & vbCrLf
'                    SqlLV = SqlLV & "WHERE A.CODCOLIGADA = '" & vCodcoligada & "' AND A.ATIVO = 'S' " & vbCrLf
'                    SqlLV = SqlLV & " " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE DESC,B.PROJETO"
'                End If
'                If FiltroGeral = "Não ativos" Then
'                    'SqlLV = "Select top " & LimiteLinhas & " a.iddesenho,a.desenho,a.revisao,b.fce,b.projeto,a.datacadastro,a.tipo,a.ativo from  tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto where a.codcoligada = '" & vCodcoligada & "' and a.ativo='N' order by b.fce desc,b.projeto"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " A.IDDESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.REVISAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATACADASTRO, " & vbCrLf
'                    SqlLV = SqlLV & " A.TIPO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBDESENHOS AS A " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBPROJETOS AS B ON " & vbCrLf
'                    SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
'                    SqlLV = SqlLV & " " & vbCrLf
'                    SqlLV = SqlLV & "WHERE A.CODCOLIGADA = '" & vCodcoligada & "' AND A.ATIVO = 'N' " & vbCrLf
'                    SqlLV = SqlLV & " " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE DESC,B.PROJETO"
'                End If
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Identificador", "Desenho", "Rev.", "FCE", "Projeto", "Data Cadastro", "Tipo", "Ativo", "Tipo FCE", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Cadastro de Desenhos"
'            MontaCabecalhoLV
'            MontaDadosLV "S" ' Zero a direita na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "N", "", "N", "N", "N", "E"
'                PersonaColLV 7, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 8, "N", "P", "", "N", "N", "N", "E" 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.cmdconsulta(6).ToolTipText = "Cancelar treinamento"
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'LM - LISTA DE MATERIAIS
'        ElseIf QualLV = 8 Then
'
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'
'            Set chamaForm = New frmLM
'            Formulario = "LM"
'            LegendaExc = "LM - Lista de Material" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 2 'Com quantas colunas que a varglobal irá trabalhar
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                'If MeuLV.Visible = True Then Unload MeuLV
''ESTA OCORRENDO PROBLEMAS DEVIDO A MENSAGEM DE GRAVAÇÃO AO SAIR DO FORMULARIO 'FRMLM'
'                carregaTABS "tblm", "tbfce", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then
'                    'SqlLV = "select right('0000' + rtrim(a.fce),4),a.codlm,a.dataabertura,a.descricao,a.ativo,CASE WHEN B.status = 0 THEN 'ANDAMENTO' WHEN B.status = 1 THEN 'CONCLUIDA' WHEN B.status = 2 THEN 'PARALIZADA' END AS STATUS from tblm as a inner join tbFCE as b on a.fce = b.fce order by a.fce, a.codlm"
'                    SqlLV = ""
'                    'SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & "SELECT TOP 500 " & vbCrLf
'                    SqlLV = SqlLV & " RIGHT('0000' + RTRIM(A.FCE),4) AS FCE, " & vbCrLf
'                    SqlLV = SqlLV & " A.CODLM AS LM, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATAABERTURA AS DATAABERTURA, " & vbCrLf
'                    SqlLV = SqlLV & " A.DESCRICAO AS DESCRICAO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO AS ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 0 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 1 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 2 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
'                    SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
'                    SqlLV = SqlLV & "         A.TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "     ELSE " & vbCrLf
'                    SqlLV = SqlLV & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
'                    SqlLV = SqlLV & " END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBLM AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON A.FCE = B.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON A.FCE = FILTRO.FCE " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY A.FCE DESC, A.CODLM DESC"
'                End If
'                If FiltroGeral = "Ativos" Then
'                    'SqlLV = "select right('0000' + rtrim(a.fce),4),a.codlm,a.dataabertura,a.descricao,a.ativo,CASE WHEN B.status = 0 THEN 'ANDAMENTO' WHEN B.status = 1 THEN 'CONCLUIDA' WHEN B.status = 2 THEN 'PARALIZADA' END AS STATUS from tblm as a inner join tbFCE as b on a.fce = b.fce where a.ativo = 'S' order by a.fce, a.codlm"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " RIGHT('0000' + RTRIM(A.FCE),4) AS FCE, " & vbCrLf
'                    SqlLV = SqlLV & " A.CODLM AS LM, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATAABERTURA AS DATAABERTURA, " & vbCrLf
'                    SqlLV = SqlLV & " A.DESCRICAO AS DESCRICAO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO AS ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 0 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 1 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 2 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
'                    SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
'                    SqlLV = SqlLV & "         A.TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "     ELSE " & vbCrLf
'                    SqlLV = SqlLV & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
'                    SqlLV = SqlLV & " END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBLM AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON A.FCE = B.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON A.FCE = FILTRO.FCE WHERE A.ATIVO = 'S' " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY A.FCE DESC, A.CODLM DESC"
'                End If
'                If FiltroGeral = "Não ativos" Then
'                    'SqlLV = "select right('0000' + rtrim(a.fce),4),a.codlm,a.dataabertura,a.descricao,a.ativo,CASE WHEN B.status = 0 THEN 'ANDAMENTO' WHEN B.status = 1 THEN 'CONCLUIDA' WHEN B.status = 2 THEN 'PARALIZADA' END AS STATUS from tblm as a inner join tbFCE as b on a.fce = b.fce where a.ativo <> 'S' order by a.fce, a.codlm"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " RIGHT('0000' + RTRIM(A.FCE),4) AS FCE, " & vbCrLf
'                    SqlLV = SqlLV & " A.CODLM AS LM, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATAABERTURA AS DATAABERTURA, " & vbCrLf
'                    SqlLV = SqlLV & " A.DESCRICAO AS DESCRICAO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO AS ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 0 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'ANDAMENTO' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 1 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'CONCLUIDA' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN B.STATUS = 2 THEN " & vbCrLf
'                    SqlLV = SqlLV & "         'PARALIZADA' " & vbCrLf
'                    SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
'                    SqlLV = SqlLV & "         A.TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "     ELSE " & vbCrLf
'                    SqlLV = SqlLV & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
'                    SqlLV = SqlLV & " END AS TIPO_FCE " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBLM AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS B ON A.FCE = B.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBPEDIDOS AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON A.FCE = FILTRO.FCE WHERE A.ATIVO <> 'S' " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY A.FCE DESC, A.CODLM DESC"
'                End If
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "FCE", "LM", "Data Abertura", "Descrição", "Ativo", "Status FCE", "Tipo FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "LM's - Listas de Materiais"
'            MontaCabecalhoLV
'            MontaDadosLV "N"
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "S", "", "N", "S", "N", "D"
'                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 5, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 6, "N", "P", "", "N", "N", "N", "E" 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.cmdconsulta(5).ToolTipText = "Editar LM - Lista de Materiais"
'            MeuLV.Visible = True
'            Exit Sub
'        'MP - Métodos e Processos
'        ElseIf QualLV = 9 Then
'
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmMPCompleto
'
'
'            Formulario = "Métodos & Processos"
'            LegendaExc = "Método & Processo" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1 'Com quantas colunas que a varglobal irá trabalhar
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                'If MeuLV.Visible = True Then
'                '    Unload MeuLV
'                'End If
'                carregaTABS "tbMP", "tbProjetos", "tbMPItens", "tbitemlm", "tbdesenhos", "tbos", "tbretrabalho", "tcfce", "", ""
'
'                If FiltroGeral = "Todos" Then
'                    'SqlLV = "select top " & LimiteLinhas & " a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produção' when a.status = 3 then 'Expedição' else 'Planejamento' end as status,CASE WHEN h.status = 0 THEN 'ANDAMENTO' WHEN h.status = 1 THEN 'CONCLUIDA' WHEN h.status = 2 THEN 'PARALIZADA' END AS Status_FCE " & _
'                    '                                  "from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao " & _
'                    '                                  "inner join tbFCE as h on b.fce = h.fce where a.ativo = 'S' group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status,h.status order by a.idprogramacao desc"
'                    SqlLV = ""
'                    SqlLV = SqlLV & "SELECT TOP " & LimiteLinhas & " " & vbCrLf
'                    SqlLV = SqlLV & " A.IDPROGRAMACAO, " & vbCrLf
'                    SqlLV = SqlLV & " C.IDOS, " & vbCrLf
'                    SqlLV = SqlLV & " F.REVISAO, " & vbCrLf
'                    SqlLV = SqlLV & " A.DATAPROGRAMACAO, " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE, " & vbCrLf
'                    SqlLV = SqlLV & " B.PROJETO, " & vbCrLf
'                    SqlLV = SqlLV & " A.RESPONSAVEL, " & vbCrLf
'                    SqlLV = SqlLV & " MIN(E.DESENHO) AS DESENHO, " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO, " & vbCrLf
'                    SqlLV = SqlLV & " MAX(G.IDRETRABALHO) AS RETRABALHO, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN A.STATUS = 1 THEN 'Planejamento' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN A.STATUS > 1 AND A.STATUS < 3 THEN 'Produção' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN A.STATUS = 3 THEN 'Expedição' " & vbCrLf
'                    SqlLV = SqlLV & "     ELSE 'Planejamento' " & vbCrLf
'                    SqlLV = SqlLV & " END AS STATUS, " & vbCrLf
'                    SqlLV = SqlLV & " CASE " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN H.STATUS = 0 THEN 'ANDAMENTO' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN H.STATUS = 1 THEN 'CONCLUIDA' " & vbCrLf
'                    SqlLV = SqlLV & "     WHEN H.STATUS = 2 THEN 'PARALIZADA' " & vbCrLf
'                    SqlLV = SqlLV & " END AS STATUS_FCE, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END AS TIPO_FCE, " & vbCrLf
'                    SqlLV = SqlLV & " CASE WHEN F.TIPOOS = 0 THEN 'Fabricação' WHEN F.TIPOOS = 1 THEN 'Manutenção' WHEN F.TIPOOS = 2 THEN 'Usinagem' ELSE 'Fabricação' END AS TIPO " & vbCrLf
'                    SqlLV = SqlLV & "FROM TBMP AS A " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBPROJETOS AS B ON " & vbCrLf
'                    SqlLV = SqlLV & " A.CODPROJETO = B.CODPROJETO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBMPITENS AS C ON " & vbCrLf
'                    SqlLV = SqlLV & " A.IDPROGRAMACAO = C.IDPROGRAMACAO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBITEMLM AS D ON " & vbCrLf
'                    SqlLV = SqlLV & " SUBSTRING(C.DESENHOS,1,2) = D.CODLM AND " & vbCrLf
'                    SqlLV = SqlLV & " REPLACE(SUBSTRING(C.DESENHOS,3,4),';','') = D.CODSEQ AND " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE = D.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBDESENHOS AS E ON " & vbCrLf
'                    SqlLV = SqlLV & " D.CODIGODES = E.IDDESENHO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBOS AS F ON " & vbCrLf
'                    SqlLV = SqlLV & " C.IDOS = F.IDOS AND " & vbCrLf
'                    SqlLV = SqlLV & " C.REVISAOOS = F.REVISAO " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN TBRETRABALHO AS G ON " & vbCrLf
'                    SqlLV = SqlLV & " A.IDPROGRAMACAO = G.IDPROGRAMACAO " & vbCrLf
'                    SqlLV = SqlLV & "INNER JOIN TBFCE AS H ON " & vbCrLf
'                    SqlLV = SqlLV & " B.FCE = H.FCE " & vbCrLf
'                    SqlLV = SqlLV & "LEFT JOIN " & vbCrLf
'                    SqlLV = SqlLV & " ( " & vbCrLf
'                    SqlLV = SqlLV & " SELECT  FCE, " & vbCrLf
'                    SqlLV = SqlLV & "     COALESCE( " & vbCrLf
'                    SqlLV = SqlLV & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
'                    SqlLV = SqlLV & "          FROM TBPEDIDOS AS O " & vbCrLf
'                    SqlLV = SqlLV & "          WHERE O.FCE  = C.FCE " & vbCrLf
'                    SqlLV = SqlLV & "          GROUP BY TIPOFCEDESC " & vbCrLf
'                    SqlLV = SqlLV & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
'                    SqlLV = SqlLV & " FROM TBLM AS C " & vbCrLf
'                    SqlLV = SqlLV & " GROUP BY FCE " & vbCrLf
'                    SqlLV = SqlLV & " ) AS FILTRO ON B.FCE = FILTRO.FCE " & vbCrLf
'                    SqlLV = SqlLV & "WHERE " & vbCrLf
'                    SqlLV = SqlLV & " A.ATIVO = 'S' " & vbCrLf
'                    SqlLV = SqlLV & "GROUP BY A.IDPROGRAMACAO,C.IDOS,F.REVISAO,A.DATAPROGRAMACAO,B.FCE,B.PROJETO,A.RESPONSAVEL,A.ATIVO,A.STATUS,H.STATUS,FILTRO.TIPO,F.TIPOOS " & vbCrLf
'                    SqlLV = SqlLV & "ORDER BY A.IDPROGRAMACAO DESC"
'                End If
'
''                If FiltroGeral = "Planejamento" Then SqlLV = "select a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produção' when a.status = 3 then 'Expedição' else 'Planejamento' end as status,CASE WHEN h.status = 0 THEN 'ANDAMENTO' WHEN h.status = 1 THEN 'CONCLUIDA' WHEN h.status = 2 THEN 'PARALIZADA' END AS Status_FCE " & _
''                                                             "from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos  and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao " & _
''                                                             "inner join tbFCE as h on b.fce = h.fce where a.status = 1 group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status,h.status order by a.idprogramacao"
'
''                If FiltroGeral = "Produção" Then SqlLV = "select a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produção' when a.status = 3 then 'Expedição' else 'Planejamento' end as status,CASE WHEN h.status = 0 THEN 'ANDAMENTO' WHEN h.status = 1 THEN 'CONCLUIDA' WHEN h.status = 2 THEN 'PARALIZADA' END AS Status_FCE " & _
''                                                         "from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao " & _
''                                                         "inner join tbFCE as h on b.fce = h.fce where a.status > 1 and a.status < 3 group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status,H.STATUS order by a.idprogramacao"
'
''                If FiltroGeral = "Expedição" Then SqlLV = "select a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produção' when a.status = 3 then 'Expedição' else 'Planejamento' end as status, CASE WHEN h.status = 0 THEN 'ANDAMENTO' WHEN h.status = 1 THEN 'CONCLUIDA' WHEN h.status = 2 THEN 'PARALIZADA' END AS Status_FCE " & _
''                                                          "from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos  and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao " & _
''                                                          "inner join tbFCE as h on b.fce = h.fce where a.status = 3 group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status,H.STATUS order by a.idprogramacao"
'
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            'MeuLV.ListView1.CheckBoxes = True
'            'MeuLV.cmdconsulta(0).Visible = True
'            'MeuLV.cmdconsulta(9).Visible = True
'            'MeuLV.cmdconsulta(11).Visible = True
'            'MeuLV.cmdconsulta(12).Visible = True
'            QtdColReal = 0
'            MontaCabLV "Planejamento", "OS nº", "Rev.", "Data", "FCE", "Projeto", "Responsável", "Desenho", "Ativo", "Retrabalho", "Status", "Status FCE", "Tipo FCE", "Tipo OS", "", "", "", "", "", "", ""
'            DimensionaLV "Métodos e Processos"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                PersonaColLV 1, "N", "N", "", "N", "S", "N", "E"
'                PersonaColLV 8, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 9, "S", "N", "", "N", "N", "N", "E"
'                PersonaColLV 10, "N", "S", "", "N", "N", "N", "E"
'                PersonaColLV 11, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 12, "N", "P", "", "N", "N", "N", "E" 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(9).PictureNormal = MeuLV.ImageList1.ListImages(16).Picture
'            MeuLV.cmdconsulta(9).ToolTipText = "CD - Comunicação de Desvio"
'
'            Set MeuLV.cmdconsulta(11).PictureNormal = MeuLV.ImageList1.ListImages(9).Picture
'            MeuLV.cmdconsulta(11).ToolTipText = "Abertura de Retrabalho"
'
'            Set MeuLV.cmdconsulta(12).PictureNormal = MeuLV.ImageList1.ListImages(14).Picture
'            MeuLV.cmdconsulta(12).ToolTipText = "Baixa Parcial de OS/Operação"
'
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(0).Visible = True
'            MeuLV.cmdconsulta(9).Visible = True
'            MeuLV.cmdconsulta(11).Visible = True
'            MeuLV.cmdconsulta(12).Visible = True
'
'            MeuLV.Visible = True
'            Exit Sub
''        'Controle de Desenhos
'        ElseIf QualLV = 10 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmCD
'            Formulario = "CD - Controle de Desenhos"
'            LegendaExc = "CD - Controle de Desenhos" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                frmFiltro.frmPeriodo.Visible = True
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                If MeuLV.Visible = True Then Unload MeuLV
'                carregaTABS "tbcd", "tbdesenhos", "tbprojetos", "", "", "", "", "", "", ""
'
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,CAST(c.fce AS VARCHAR(4)) + ' - ' + c.projeto AS FCE,b.desenho,b.revisao,a.quantidade,a.pesounit,(a.quantidade*a.pesounit) as pesototal,a.datarecebido,a.ptempo + ' ' + a.punidade,a.usuario,a.dataini,a.datafim,a.croqui,a.status,a.observacao,a.ativo,a.detalhista from " & _
'                                                      "tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto order by a.idcd desc"
''                If FiltroGeral = "Ativos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,CAST(c.fce AS VARCHAR(4)) + ' - ' + c.projeto AS FCE,b.desenho,b.revisao,a.quantidade,a.pesounit,(a.quantidade*a.pesounit) as pesototal,a.datarecebido,a.ptempo + ' ' + a.punidade,a.usuario,a.dataini,a.datafim,a.croqui,a.status,a.observacao,a.ativo,a.detalhista from " & _
''                                                      "tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto where a.ativo = 'S' order by a.idcd desc"
''                If FiltroGeral = "Não ativos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,CAST(c.fce AS VARCHAR(4)) + ' - ' + c.projeto AS FCE,b.desenho,b.revisao,a.quantidade,a.pesounit,(a.quantidade*a.pesounit) as pesototal,a.datarecebido,a.ptempo + ' ' + a.punidade,a.usuario,a.dataini,a.datafim,a.croqui,a.status,a.observacao,a.ativo,a.detalhista from " & _
''                                                      "tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto where a.ativo <> 'S' order by a.idcd desc"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Identificador", "FCE", "Desenho", "Rev.", "Quant.", "Peso Unit.", "Peso Total", "Recebido", "Previsão Det.", "Usuário", "Data inicio", "Data fim", "Croqui", "Status", "Observação", "Ativo", "Detalhista", "", "", "", ""
'            DimensionaLV "CD - Controle de Desenhos"
'            MontaCabecalhoLV
'            MontaDadosLV "S" ' Zero a direita na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "N", "", "N", "N", "N", "E"
'                PersonaColLV 4, "N", "N", "", "N", "N", "N", "D"
'                PersonaColLV 5, "N", "N", "", "N", "N", "S", "D"
'                PersonaColLV 6, "N", "N", "", "N", "N", "S", "D"
'
'                PersonaColLV 13, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 15, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.cmdconsulta(6).ToolTipText = "Cancelar treinamento"
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
''        'Fórmula = Centro de Custo
'        ElseIf QualLV = 11 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmFormulaCC
'            Formulario = "Fórmula - Centro de Custo"
'            LegendaExc = "Fórmula - Centro de Custo" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "GCCUSTO", "tbFormula", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select a.CODREDUZIDO,a.NOME, 'formula' = case when max(b.nmform) IS NULL then '-' else 'com formula' end from " & vBancoTotvs & ".dbo.GCCUSTO as a left join " & sDatabaseName & ".dbo.tbFormula as b " & _
'                "on a.CODREDUZIDO = b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS where a.codcoligada = '" & vCodcoligada & "' and (ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '1000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '3000' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '7000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '5000' or " & _
'                "ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '6000' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '9001' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '7000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '4000' or ativo  = 'T') group by a.ID,a.CODREDUZIDO,a.NOME order by a.CODREDUZIDO"
''                If FiltroGeral = "Ativos" Then SqlLV = "select a.codavaliacao,a.nomeavaliacao,a.tipo,a.peso,a.ativo from tbAvaliacao as a where a.codcoligada = '" & vCodcoligada & "' and a.ativo = 'S'"
''                If FiltroGeral = "Não ativos" Then SqlLV = "select a.codavaliacao,a.nomeavaliacao,a.tipo,a.peso,a.ativo from tbAvaliacao as a where a.codcoligada = '" & vCodcoligada & "' and a.ativo is null or a.codcoligada = '" & vCodcoligada & "' and ativo ='N'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Centro de Custo", "Nome Centro de Custo", "Fórmula", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Fórmulas - Centro de Custo"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                'PersonaColLV 3, "N", "N", "", "N", "N", "N", "D"
'                'PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'QUALIDADE - RNCF (Registro de Não Conformidade de Fabricação)
'        ElseIf QualLV = 12 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmRNCF
'            Formulario = "RNCF - Registro de Não Conformidade de Fabricação"
'            LegendaExc = "RNCF - Registro de Não Conformidade de Fabricação" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbComunicacaoDesvio", "tbMPItens", "tbMP", "tbProjetos", "tbRNC", "tbRetrabalho", "", "", "", ""
'
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                      "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd where a.idcd >= 1 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                'If FiltroGeral = "CD - Comunicação de Desvio" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                '                                          "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                '                                          "where a.status = 6 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                'If FiltroGeral = "CODAC - Coleta de Dados e Análise de Causas" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                '                                          "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                '                                          "where a.status = 7 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                'If FiltroGeral = "DAAC - Definição da Ação e Análise Concluida" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                '                                           "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                '                                           "where a.status = 8 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'
'                'If FiltroGeral = "EVA - Execução e Verificação da Ação" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                '                                           "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                '                                           "where a.status = 9 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                'If FiltroGeral = "TAC - Tomada de Ação Concluida" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                '                                           "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                '                                           "where a.status = 10 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = True
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "CD nº", "Data Abertura", "Responsável", "OS nº", "FCE", "Projeto", "Observação", "Status", "RNC nº", "Data Conclusão", "Retrabalho", "Retrabalho nº", "Data Fechamento", "", "", "", "", "", "", "", ""
'            DimensionaLV "RNCF - Registro de Não Conformidade de Fabricação"
'            MontaCabecalhoLV
'            MontaDadosLV "S" ' Zero a direita na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 3, "N", "N", "", "N", "S", "N", "E"
'                'PersonaColLV 4, "N", "N", "", "N", "S", "N", "E"
'                'PersonaColLV 6, "N", "N", "", "N", "S", "N", "E"
'                PersonaColLV 7, "S", "S", "", "S", "N", "N", "E"
'                PersonaColLV 8, "S", "N", "", "N", "S", "N", "E"
'                PersonaColLV 10, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 11, "S", "S", "", "N", "S", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(9).PictureNormal = MeuLV.ImageList1.ListImages(18).Picture
'            MeuLV.cmdconsulta(9).ToolTipText = "Causais"
'            MeuLV.Visible = True
'            Exit Sub
'        'USUÁRIOS
'        ElseIf QualLV = 13 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmUsuarios
'            Formulario = "Usuários"
'            LegendaExc = "Usuário" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                carregaTABS "tbusuarios", "tbgrupo", "", "", "", "", "", "", "", ""
'
'                If FiltroGeral = "Todos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where b.codcoligada = " & vCodcoligada
'                'If FiltroGeral = "Ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo = 'S'"
'                'If FiltroGeral = "Não ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo is null or a.ativo ='N'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Código", "Nome do usuário", "Grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Usuários"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'GRUPOS
'        ElseIf QualLV = 14 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmGrupos
'            Formulario = "Grupos"
'            LegendaExc = "Grupo" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                carregaTABS "tbGrupo", "", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select a.codigo,a.descricao,a.ativo from tbgrupo as a"
'                'If FiltroGeral = "Ativos" Then SqlLV = "select a.codigo,a.descricao,a.ativo from tbgrupo where ativo = 'S'"
'                'If FiltroGeral = "Não ativos" Then SqlLV = "select a.codigo,a.descricao,a.ativo from tbgrupo where ativo is null or ativo ='N'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Código", "Nome do grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Grupos"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'OS FECHAMENTO - PERMISSÃO DE COLABORADORES
'        ElseIf QualLV = 15 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'
'            Set chamaForm = New frmPerColab
'            Formulario = "OS Fechamento - Permissão de Colaboradores"
'            LegendaExc = "Permissão do colaborador" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                If MeuLV.Visible = True Then Unload MeuLV
'                carregaTABS "PFUNC", "PPESSOA", "tbautfechaos", "", "", "", "", "", "", ""
'
''                If FiltroGeral = "Todos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
''                                                      "a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 left join tbautfechaos as c on a.chapa = c.chapa COLLATE SQL_Latin1_General_CP1_CI_AS where a.CHAPA > 0 ORDER BY a.chapa"
'
'                If FiltroGeral = "Todos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from " & vBancoTotvs & ".dbo.PFUNC as a inner join " & vBancoTotvs & ".dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 " & _
'                                                      "left join tbautfechaos as c on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI = c.chapa where a.CHAPA > 0 union " & _
'                                                      "select a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI,a.nome COLLATE SQL_Latin1_General_CP1_CI_AI,case when b.chapa is not null then 'S' else 'N' end as ativo from tbTerceirizados as a left join tbautfechaos as b on a.chapa = b.chapa and a.ativo = 'S' ORDER BY a.chapa"
'
''                If FiltroGeral = "Ativos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
''                                                       "a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 left join tbautfechaos as c on a.chapa = c.chapa COLLATE SQL_Latin1_General_CP1_CI_AS where  c.chapa is not null ORDER BY a.chapa"
''                If FiltroGeral = "Não ativos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
''                                                           "a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 left join tbautfechaos as c on a.chapa = c.chapa COLLATE SQL_Latin1_General_CP1_CI_AS where  c.chapa is null or c.chapa = 'N' ORDER BY a.chapa"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
''
'            QtdColReal = 0
'            MontaCabLV "Chapa", "Nome", "Permissão", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "OS Fechamento - Permissão de Colaboradores"
'            MontaCabecalhoLV
'            MontaDadosLV "N" 'Coloca zeros a esquerda na primeira coluna
'            If checaFiltro = True Then
''                PersonaColLV 4, "S", "S", "%", "N", "N", "S", "D"
'                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'LF - Liberação de Fabricação
'        ElseIf QualLV = 16 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmRelInsp
'            Formulario = "Relatórios de Inspeção"
'            LegendaExc = "Relatórios de Inspeção" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                carregaTABS "tbProjetos", "tbFO", "tbCliFor", "tbFCE", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.codprojeto,a.fce,a.projeto,c.nome,CASE WHEN d.status = 0 THEN 'ANDAMENTO' WHEN d.status = 1 THEN 'CONCLUIDA' WHEN d.status IS NULL THEN 'DUVIDA' WHEN d.status = 2 THEN 'PARALIZADA' END AS STATUS from tbProjetos as a inner join tbFO as b on a.fce=b.fce inner join tbclifor as c on b.codclifor = c.codclifor inner join tbFCE as d on b.fce = d.fce where a.fce > 2000 Order by a.fce desc,a.descricao"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(5).Visible = False
'            MeuLV.cmdconsulta(6).Visible = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'
'            QtdColReal = 0
'            MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "Nome Cliente", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Relatórios de Inspeção"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
''                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
'
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(19).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Relatório de Inspeção - Fabricação"
'            Set MeuLV.cmdconsulta(6).PictureNormal = MeuLV.ImageList1.ListImages(21).Picture
'            MeuLV.cmdconsulta(6).ToolTipText = "Relatório de Inspeção - Pintura"
'            MeuLV.Visible = True
'
'            Exit Sub
'        'RO - Relatório de Expedição
'        ElseIf QualLV = 17 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = frmRelExp
'            Formulario = "Relatórios de Expedição"
'            LegendaExc = "Relatórios de Expedição" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbProjetos", "tbFO", "tbCliFor", "tbFCE", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.codprojeto,a.fce,a.projeto,c.nome,CASE WHEN d.status = 0 THEN 'ANDAMENTO' WHEN d.status = 1 THEN 'CONCLUIDA' WHEN d.status IS NULL THEN 'DUVIDA' WHEN d.status = 2 THEN 'PARALIZADA' END AS STATUS from tbProjetos as a inner join tbFO as b on a.fce=b.fce inner join tbclifor as c on b.codclifor = c.codclifor inner join tbFCE as d on b.fce = d.fce where a.fce > 2000 Order by a.fce desc,a.descricao"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(5).Visible = True
'            MeuLV.cmdconsulta(6).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'
'            QtdColReal = 0
'            MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "Nome Cliente", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Relatórios de Expedição"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
''                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(28).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Relatório de Expedição"
'            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(27).Picture
'            MeuLV.cmdconsulta(5).ToolTipText = "Relatório de Expedição Avulso"
'            MeuLV.Visible = True
'            Exit Sub
'        'IMPRESSAO DOS RELATÓRIOS DE EXPEDIÇÃO
'        ElseIf QualLV = 18 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New FCRExpedicao
'            Formulario = "Relatórios de Expedição emitidos"
'            LegendaExc = "Relatório de Expedição" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                carregaTABS "tbRelInspExp", "rbProjetos", "tbFCE", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.codrel,case when a.fce = 0 then NULL else a.fce end as FCE,b.projeto,b.descricao,a.datarel,case when a.statusimp = 0 then 'Não impresso' else 'Impresso' end,CASE WHEN c.status = 0 THEN 'ANDAMENTO' WHEN c.status = 1 THEN 'CONCLUIDA' WHEN c.status IS NULL THEN 'DUVIDA' WHEN c.status = 2 THEN 'PARALIZADA' END AS STATUS from tbRelInspExp as a left join tbProjetos as b on a.fce = b.fce and a.codprojeto = b.codprojeto left join tbFCE as c on b.fce = c.fce where a.tiporel = 11"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(5).Visible = False
'            MeuLV.cmdconsulta(6).Visible = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'
'            QtdColReal = 0
'            MontaCabLV "Nº Relatório", "FCE", "Projeto", "Descrição", "Data emissão", "Status Impressão", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Impressão - Relatórios de Expedição"
'            MontaCabecalhoLV
'            MontaDadosLV "S" 'Coloca zeros a esquerda na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "S", "", "N", "N", "N", "E"
'                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'IMPRESSAO DOS RELATÓRIOS DE INSPEÇÃO (QUALIDADE)
'        ElseIf QualLV = 19 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'
'            Set chamaForm = New FCRLibFab
'            Formulario = "Relatórios de Inspeção emitidos"
'            LegendaExc = "Relatório de Inspeção" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbRelInspExp", "tbProjetos", "tbFCE", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.codrel,a.fce,b.projeto,b.descricao, a.datarel,case when a.statusimp = 0 then 'Não impresso' else 'Impresso' end,case when a.tiporel = 3 then 'FABRICAÇÃO' else 'PINTURA' end,CASE WHEN c.status = 0 THEN 'ANDAMENTO' WHEN c.status = 1 THEN 'CONCLUIDA' WHEN c.status IS NULL THEN 'DUVIDA' WHEN c.status = 2 THEN 'PARALIZADA' END AS STATUS from tbRelInspExp as a inner join tbProjetos as b on a.fce = b.fce and a.codprojeto = b.codprojeto left join tbFCE as c on b.fce = c.fce where a.tiporel < 11"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(5).Visible = False
'            MeuLV.cmdconsulta(6).Visible = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'
'            QtdColReal = 0
'            MontaCabLV "Nº Relatório", "FCE", "Projeto", "Descrição", "Data emissão", "Status Impressão", "Tipo", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Impressão - Relatórios de Inspeção"
'            MontaCabecalhoLV
'            MontaDadosLV "S" 'Coloca zeros a esquerda na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "S", "", "N", "N", "N", "E"
'                PersonaColLV 7, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        'FATURAMENTO POR FCE
'        ElseIf QualLV = 20 Then
'
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'
'            Set chamaForm = New FCRFatFCE
'            Formulario = "Faturamento por FCE"
'            LegendaExc = "Faturamento" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'
'                carregaTABS "", "", "", "", "", "", "", "", "", ""
'
'                If FiltroGeral = "Em aberto" Then SqlLV = "SELECT T1.DESCRICAO,T1.CODTB3FAT,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO-isnull(T4.DEVOLVIDO,0) VALOR_BRUTO,T1.VALOR_LIQUIDO-isnull(T4.DEVOLVIDO,0) VALOR_LIQUIDO,T1.DTCRIACAO,(T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL,T2.VALOR_BAIXADO as VALOR_BAIXADO, " & _
'                                                      "((T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-isnull(T5.ADIANTADO,0)-isnull(T6.CANCELADO,0) AS VALOR_RECEBER,T3.PESO,T3.VALOR_TOTAL,CASE WHEN T1.status = 0 THEN 'ANDAMENTO' WHEN T1.status = 1 THEN 'CONCLUIDA' WHEN T1.status IS NULL THEN 'DUVIDA' WHEN T1.status = 2 THEN 'PARALIZADA' END AS STATUS FROM (select MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO, " & _
'                                                      "SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO,MAX(c.status) as status FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' left join tbFCE as C on a.CODTB3FAT = c.fce GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) " & _
'                                                      ") T1 LEFT JOIN (SELECT B.CODTB3FAT,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY B.CODTB3FAT) T2 " & _
'                                                      "ON T1.CODTB3FAT = T2.CODTB3FAT LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from " & sDatabaseName & ".DBO.tbPedidos AS A group by a.fce) T3 ON T1.CODTB3FAT = T3.FCE LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as DEVOLVIDO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT " & _
'                                                      "where B.CODTMV in ('1.2.15','1.2.17') and B.STATUS = 'F' group by B.CODTB3FAT) T4 ON T1.CODTB3FAT = T4.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as ADIANTADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.25') group by B.CODTB3FAT) T5 ON T1.CODTB3FAT = T5.CODTB3FAT " & _
'                                                      "LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as CANCELADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT where B.STATUS = 'C' and B.CODTMV in ('2.2.01','2.2.05','1.2.15','1.2.17') group by b.CODTB3FAT) T6 ON T1.CODTB3FAT = T6.CODTB3FAT where T2.VALOR_RECEBER > 0 or T2.VALOR_RECEBER is null ORDER BY T1.CODTB3FAT"
'                If FiltroGeral = "Todos" Then SqlLV = "SELECT T1.DESCRICAO,T1.CODTB3FAT,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO-isnull(T4.DEVOLVIDO,0) VALOR_BRUTO,T1.VALOR_LIQUIDO-isnull(T4.DEVOLVIDO,0) VALOR_LIQUIDO,T1.DTCRIACAO,(T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL,T2.VALOR_BAIXADO as VALOR_BAIXADO, " & _
'                                                      "((T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-isnull(T5.ADIANTADO,0)-isnull(T6.CANCELADO,0) AS VALOR_RECEBER,T3.PESO,T3.VALOR_TOTAL,CASE WHEN T1.status = 0 THEN 'ANDAMENTO' WHEN T1.status = 1 THEN 'CONCLUIDA' WHEN T1.status IS NULL THEN 'DUVIDA' WHEN T1.status = 2 THEN 'PARALIZADA' END AS STATUS FROM (select MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO, " & _
'                                                      "SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO,MAX(c.status) as status FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' left join tbFCE as C on a.CODTB3FAT = c.fce GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) " & _
'                                                      ") T1 LEFT JOIN (SELECT B.CODTB3FAT,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY B.CODTB3FAT) T2 " & _
'                                                      "ON T1.CODTB3FAT = T2.CODTB3FAT LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from " & sDatabaseName & ".DBO.tbPedidos AS A group by a.fce) T3 ON T1.CODTB3FAT = T3.FCE LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as DEVOLVIDO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT " & _
'                                                      "where B.CODTMV in ('1.2.15','1.2.17') and B.STATUS = 'F' group by B.CODTB3FAT) T4 ON T1.CODTB3FAT = T4.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as ADIANTADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.25') group by B.CODTB3FAT) T5 ON T1.CODTB3FAT = T5.CODTB3FAT " & _
'                                                      "LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as CANCELADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT where B.STATUS = 'C' and B.CODTMV in ('2.2.01','2.2.05','1.2.15','1.2.17') group by b.CODTB3FAT) T6 ON T1.CODTB3FAT = T6.CODTB3FAT ORDER BY T1.CODTB3FAT"
'
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(5).Visible = False
'            MeuLV.cmdconsulta(6).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            MeuLV.cmdconsulta(0).Visible = True
'            QtdColReal = 0
'            MontaCabLV "Descrição", "FCE", "Peso Líquido (FAT)", "Peso Bruto (FAT)", "Valor Bruto (FAT)", "Valor Líquido (FAT)", "Data Cadastro(FCE)", "Valor Original (FIN)", "Valor Baixado (FIN)", "Valor Receber (FIN)", "Peso (COM)", "Valor Vendido (COM)", "Status", "", "", "", "", "", "", "", ""
'            DimensionaLV "Faturamento por FCE"
'            MontaCabecalhoLV
'            MontaDadosLV "N" 'Coloca zeros a esquerda na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "S", "", "N", "N", "N", "E"
'                PersonaColLV 2, "N", "N", "", "N", "N", "S", "D"
'                PersonaColLV 3, "N", "N", "", "N", "N", "S", "D"
'                PersonaColLV 4, "N", "N", "R$ ", "N", "N", "S", "D"
'                PersonaColLV 5, "N", "N", "R$ ", "N", "N", "S", "D"
'                PersonaColLV 7, "N", "N", "R$ ", "N", "N", "S", "D"
'                PersonaColLV 8, "N", "N", "R$ ", "N", "N", "S", "D"
'                PersonaColLV 9, "S", "S", "R$ ", "N", "N", "S", "D"
'                PersonaColLV 10, "N", "N", "", "N", "N", "S", "D"
'                PersonaColLV 11, "S", "S", "R$ ", "N", "N", "S", "D"
'                PersonaColLV 12, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Set MeuLV.cmdconsulta(0).PictureNormal = MeuLV.ImageList1.ListImages(29).Picture
'            MeuLV.cmdconsulta(0).ToolTipText = "Alterar Status da FCE"
'            MeuLV.Visible = True
'            Exit Sub
'
'        'TERCEIRIZADOS
'        ElseIf QualLV = 21 Then
'            If Pesquisa <> "filtro" Then
'                MeuLV.Visible = False
'                'frmMsgAutomatica.Show 1
'            End If
'            Set chamaForm = New frmTerceirizados
'            Formulario = "Terceirizados"
'            LegendaExc = "Terceirizados" 'Usado na mensagem de exclusão
'            indiceVarGlobal = 1
'            Permissao
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then
'                    frmFiltro.Show 1
'                Else
'                    filtroPadrao
'                    'LocalString1 SqlLV
'                    'SqlLV = Replace(SqlLV, vMantemExpressao, vSubstituto)
'                End If
'                carregaTABS "tbusuarios", "tbgrupo", "", "", "", "", "", "", "", ""
'                If FiltroGeral = "Todos" Then SqlLV = "select a.chapa,a.nome,a.idsetor,a.setor,a.idfuncao,a.funcao,a.idcc,a.nmcc,a.empresa,a.datacadastro,a.datacontratoini,a.datacontratofim,a.ativo from tbTerceirizados as a"
'                'If FiltroGeral = "Ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo = 'S'"
'                'If FiltroGeral = "Não ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo is null or a.ativo ='N'"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Código", "Nome do usuário", "ID Setor", "Nome Setor", "ID Função", "Nome Função", "ID CC", "Nome CC", "Empresa", "D. Cadastro", "D. Contrato ini.", "D. Contrato Fim", "Ativo", "", "", "", "", "", "", "", ""
'            DimensionaLV "Terceirizados"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                PersonaColLV 12, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            MeuLV.Visible = True
'            Exit Sub
'        End If
'
'        Set frmFiltro = Nothing
'        Set MeuLV = Nothing
'        Set chamaForm = Nothing
'TrataErro:
'    If Err.Number = 400 Then
'        FiltroGeral = "Ativos"
'        Resume Next
'    End If
'End Sub

Public Sub carregaTABS(vTab1 As String, vTab2 As String, vTab3 As String, vTab4 As String, vTab5 As String, vTab6 As String, vTab7 As String, vTab8 As String, vTab9 As String, vTab10 As String)
    vTabela1 = vTab1
    vTabela2 = vTab2
    vTabela3 = vTab3
    vTabela4 = vTab4
    vTabela5 = vTab5
    vTabela6 = vTab6
    vTabela7 = vTab7
    vTabela8 = vTab8
    vTabela9 = vTab9
    vTabela10 = vTab10
End Sub

''AINDA NÃO FOI REALIZADA ADAPTAÇÃO PARA OS COMPONENTES DINÂMICOS
'Public Sub CarregaSQLExcluir(QLV As Integer)
'On Error GoTo Err
'    Dim rsExcLVGeral As New ADODB.Recordset
'    Dim P As Integer
'    If QLV = 0 Then
'        'frmDemitirColaborador.Show 1
'        'gravaLog varGlobal, MeuLV.ListView1.SelectedItem.ListSubItems.Item(1), "-"
'    ElseIf QLV = 1 Then
'        'SqlExcLVGeral = "Delete from tbColaboradores where a.codcoligada = '" & vCodcoligada & "' and cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresesc where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradorescur where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresexp where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoreshist where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'"
'    ElseIf QLV = 2 Then
'        'SqlExcLVGeral = "Delete from tbDepartamentos where codDepartamento= '" & Val(varGlobal) & "' ;Delete from tbDepartamentosHistResp where codDepartamento= '" & Val(varGlobal) & "'"
'10      cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    SqlExcLVGeral = "UPDATE tbDepartamentos set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and coddepartamento = " & Val(MeuLV.ListView1.ListItems.Item(P))
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                End If
'            Next
'        End If
'        cnBanco.CommitTrans
'    ElseIf QLV = 3 Then
'        'SqlExcLVGeral = "Delete from tbSetores where codSetor= '" & Val(varGlobal) & "' ;Delete from tbSetoresHistResp where codSetor= '" & Val(varGlobal) & "'"
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    SqlExcLVGeral = "UPDATE tbSetores set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codsetor = " & Val(MeuLV.ListView1.ListItems.Item(P))
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                End If
'            Next
'        End If
'        cnBanco.CommitTrans
'    ElseIf QLV = 4 Then
'        'NAO EXCLUI O PRODUTO, EXCLUI OS DADOS DAS FÓRMULAS REFERENTE AO PRODUTO
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão da " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    SqlExcLVGeral = "Delete from tbmateriais where idprd = '" & Val(MeuLV.ListView1.ListItems.Item(P)) & "'"
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'
'                    SqlExcLVGeral = "Delete from tbConstantes where idprd = '" & Val(MeuLV.ListView1.ListItems.Item(P)) & "'"
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'
'                End If
'            Next
'        End If
'        cnBanco.CommitTrans
'
'    ElseIf QLV = 5 Then
'        'SqlExcLVGeral = "Delete from tbHabilidades where codHabilidade= " & Val(varGlobal)
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    SqlExcLVGeral = "UPDATE tbFO set ativo = 'N' where codfo = " & Val(MeuLV.ListView1.ListItems.Item(P))
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                End If
'            Next
'        End If
'        cnBanco.CommitTrans
'    ElseIf QLV = 6 Then
'        'SqlExcLVGeral = "Delete from tbEscolaridade where codescolaridade= " & Val(varGlobal)
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    SqlExcLVGeral = "UPDATE tbEscolaridade set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codescolaridade = " & Val(MeuLV.ListView1.ListItems.Item(P))
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                End If
'            Next
'        End If
'        cnBanco.CommitTrans
'    ElseIf QLV = 7 Then
'        SqlExcLVGeral = "Delete from tbdesenhos where codcoligada = '" & vCodcoligada & "' and iddesenho= '" & Val(varGlobal) & "' ;Delete from tbdesenhos where codcoligada = '" & vCodcoligada & "' and iddesenho= '" & Val(varGlobal) & "'"
'        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'    ElseIf QLV = 8 Then
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            SqlExcLVGeral = "Select count(*) from tbItemLM as a where a.fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and a.codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
'            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
'            If rsExcLVGeral.Fields(0) > 0 Then ' Desativa
'                rsExcLVGeral.Close
'                SqlExcLVGeral = "delete from tbItemLM where fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "Curso/treinamento DESATIVADO com sucesso", Ok, informacao, "ZEUS"
'            End If
'            rsExcLVGeral.Close
'
'            SqlExcLVGeral = "Select count(*) from tbLM as a where a.fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and a.codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
'            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
'            If rsExcLVGeral.Fields(0) > 0 Then ' Desativa
'                rsExcLVGeral.Close
'                SqlExcLVGeral = "delete from tbLM where fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "LM Excluida com sucesso", Ok, informacao, "ZEUS"
'            End If
'            'rsExcLVGeral.Close
'            Set rsExcLVGeral = Nothing
'        End If
'        cnBanco.CommitTrans
'        'rsExcLVGeral.Close
'        'SqlExcLVGeral = "Delete from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codTreinamento=  '" & Val(varGlobal) & "' ;Delete from tbTreinamentosRev where codcoligada = '" & vCodcoligada & "' and codTreinamento= '" & Val(varGlobal) & "'"
'    ElseIf QLV = 9 Then
'        Dim vPlanej As Integer, vOS As Integer
'        vPlanej = Val(Mid$(varGlobal, 1, 6))
'        vOS = Val(Mid$(varGlobal, 7, 6))
'        If vOS = 0 Then
'            SqlExcLVGeral = "Delete from tbmp where idprogramacao = '" & vPlanej & "' ;Delete from tbMPItens where idprogramacao = '" & vPlanej & "' ;Delete from tbositens where idprogramacao = '" & vPlanej & "' ;Delete from tbos where idos = 0"
'            rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'            mobjMsg.Abrir "Registro excluido com sucesso", Ok, informacao, "ZEUS"
'        Else
'            mobjMsg.Abrir "Registro não pode ser excluido", Ok, critico, "ZEUS"
'        End If
'
'    ElseIf QLV = 10 Then
'        Dim strResultado As String
'        mobjMsg.Abrir "Confirma o Cancelamento da " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    'If strResultado <> "" Then
'                        SqlExcLVGeral = "UPDATE tbCD set ativo = 'N' where idcd = '" & Val(varGlobal) & "'"
'                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                    'Else
'                    '    MsgBox "É necessário justificar o cancelamento"
'                    'End If
'                End If
'            Next
'            mobjMsg.Abrir "Cancelamento realizado!", Ok, critico, "Atenção"
'        End If
'    ElseIf QLV = 11 Then
'        'SqlExcLVGeral = "Delete from tbAvaliacao where codavaliacao= " & Val(varGlobal)
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    SqlExcLVGeral = "UPDATE tbAvaliacao set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codavaliacao = " & Val(MeuLV.ListView1.ListItems.Item(P))
'                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                End If
'            Next
'        End If
'        cnBanco.CommitTrans
'    ElseIf QLV = 15 Then
'        'ZEUS - Exclui Autorizados a Fechar OS - Ordem de Serviço
'        'Dim strResultado As String
'        mobjMsg.Abrir "Confirma o Cancelamento da " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'
'            For P = 1 To MeuLV.ListView1.ListItems.Count
'                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                    'If strResultado <> "" Then
'                        SqlExcLVGeral = "delete from tbAutCCusto where chapa = '" & MeuLV.ListView1.ListItems.Item(P) & "'"
'                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'
'                        SqlExcLVGeral = "delete from tbAutFechaOs where chapa = '" & MeuLV.ListView1.ListItems.Item(P) & "'"
'                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                End If
'            Next
'            mobjMsg.Abrir "Cancelamento realizado!", Ok, critico, "Atenção"
'        End If
'    ElseIf QLV = 16 Then
'        cnBanco.BeginTrans
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            frmExcluiINTD.Show 1
'        End If
'        cnBanco.CommitTrans
'    ElseIf QLV = 18 Or QLV = 19 Then
'        Dim rsProcuraItem As New ADODB.Recordset
'        Dim sqlProcuraItem As String
'        Dim vFCE As Integer, vCodLM As Integer, vCodSeq As Integer
'
'        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "ZEUS"
'        If Tp = 1 Then
'            Dim statusRel As Integer
'            If MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = "PINTURA" Then
'                statusRel = 10
'            ElseIf MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = "FABRICAÇÃO" Then
'                statusRel = 3
'            End If
'            If QLV = 18 Then
'                statusRel = 11
'            End If
'
'            'VERIFICA SE HA ALGUM STATUS MAIOR QUE O STATUS DO RELATORIO SELECIONADO
'            SqlExcLVGeral = "select a.codrel,a.fce,a.codlm,a.codseq,a.status from tbRelInspExpItens as a where a.codrel= '" & Val(varGlobal) & "'"
'            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
'            For P = 1 To rsExcLVGeral.RecordCount
'                vFCE = rsExcLVGeral.Fields(1)
'                vCodLM = rsExcLVGeral.Fields(2)
'                vCodSeq = rsExcLVGeral.Fields(3)
'                sqlProcuraItem = "select * from tbRelInspExpItens as a where a.fce = '" & vFCE & "' and a.codlm = '" & vCodLM & "' and a.codseq = '" & vCodSeq & "' and a.status > '" & statusRel & "'"
'                rsProcuraItem.Open sqlProcuraItem, cnBanco, adOpenKeyset, adLockReadOnly
'                If rsProcuraItem.RecordCount > 0 Then
'                    mobjMsg.Abrir "Relatório não pode ser excluido. O mesmo possui vínculo com o relatório: " & rsProcuraItem.Fields(0) & "", Ok, critico, "Atenção"
'                    rsProcuraItem.Close
'                    Exit Sub
'                End If
'                rsProcuraItem.Close
'                rsExcLVGeral.MoveNext
'            Next
'            rsExcLVGeral.Close
'            Set rsExcLVGeral = Nothing
'
'            'Exclui o relatório caso passe pelas condições acima
'            cnBanco.BeginTrans
'            'If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
'                'Exclui os itens do relatório
'                SqlExcLVGeral = "delete from tbRelInspExpItens where codrel = '" & Val(varGlobal) & "'"
'                 rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'
'                'Exclui os cabeçalho do relatório
'                SqlExcLVGeral = "delete from tbRelInspExp where codrel = '" & Val(varGlobal) & "'"
'                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
'                mobjMsg.Abrir "Relatório nº:" & Val(varGlobal) & " excluido com sucesso", Ok, informacao, "ZEUS"
'            'End If
'            cnBanco.CommitTrans
'
'        End If
'    End If
'    Exit Sub
'Err:
'    If Err.Number = -2147467259 Then
'        While reestabeleceConexao = False
'        Wend
'        GoTo 10
'    Else
'        Msgbox Err.Number & " - " & Err.Description
'        Resume
'    End If
'End Sub

'Calcula CPF
Public Function isCPF(ByVal pCPF As String) As Boolean
    Dim Conta As Integer, Soma As Integer, Resto As Integer, Passo As Integer
    isCPF = False: pCPF = Trim(pCPF)
    If Len(pCPF) <> 11 Then
        Exit Function
    End If
    For Passo = 11 To 12
        Soma = 0
        For Conta = 1 To Passo - 2
            Soma = Soma + Val(Mid(pCPF, Conta, 1)) * (Passo - Conta)
        Next
        Resto = 11 - (Soma - (Int(Soma / 11) * 11))
        If Resto = 10 Or Resto = 11 Then Resto = 0
        If Resto <> Val(Mid(pCPF, Passo - 1, 1)) Then
            Exit Function
        End If
    Next
    isCPF = True
End Function

Private Sub Permissao()
On Error GoTo Err
    Dim rsPermissao As New ADODB.Recordset
    Dim SqlPermissao As String
    SqlPermissao = "select * from tbConfGrupo where idgrupo = '" & XCodGrp & "' and nome = '" & Formulario & "'"
    rsPermissao.Open SqlPermissao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsPermissao.RecordCount > 0 Then
        If Not IsNull(rsPermissao.Fields(9)) Then vInc = rsPermissao.Fields(9)
        If Not IsNull(rsPermissao.Fields(10)) Then vEdi = rsPermissao.Fields(10)
        If Not IsNull(rsPermissao.Fields(11)) Then vExc = rsPermissao.Fields(11)
        If Not IsNull(rsPermissao.Fields(12)) Then vSal = rsPermissao.Fields(12)
        If Not IsNull(rsPermissao.Fields(13)) Then vImp = rsPermissao.Fields(13)
        If Not IsNull(rsPermissao.Fields(14)) Then vFil = rsPermissao.Fields(14)
    End If
    rsPermissao.Close
    Set rsPermissao = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
End Sub

Public Function reestabeleceConexao()
On Error GoTo Err
    reestabeleceConexao = False
    Dim vMensZeus As String
    vMensZeus = vMensZeus & "Não foi possível estabelecer uma conexão com a rede." & vbCrLf
    vMensZeus = vMensZeus & "Entre em contato com a equipe de TI responsável." & vbCrLf
    vMensZeus = vMensZeus & "" & vbCrLf
    vMensZeus = vMensZeus & "Clique SIM para tentar reconectar, NÃO para fechar a aplicação" & vbCrLf

    mobjMsg.Abrir vMensZeus, YesNo, pergunta, "ZEUS"
    If Tp = 1 Then
        If Conexao = True Then
            If ConexaoTotvs = True Then reestabeleceConexao = True Else reestabeleceConexao = False
        Else
            reestabeleceConexao = False
        End If

    Else
        End
    End If
    
'    If Msgbox("Não foi possível estabelecer uma conexao com a rede. Clique em SIM para tentar reestabelecer a conexão ou NÃO para fechar a aplicação", vbYesNo, "FALHA NA REDE") = vbYes Then
'        If Conexao = True Then
'            If ConexaoTotvs = True Then reestabeleceConexao = True Else reestabeleceConexao = False
'        Else
'            reestabeleceConexao = False
'        End If
'    Else
'        End
'    End If
    Exit Function
Err:
    If Err.Number = 91 Then
        Msgbox "Não foi identificado uma conexao de rede. Entre em contato com o suporte técnico", vbCritical, "ZEUS"
        End
    End If
End Function

'Gera Avaliação de Desempenho Profissional por colaborador
Public Function carregaADP()
On Error GoTo Err
    Dim rsADP As New ADODB.Recordset
    Dim sqlADP As String
    Dim X As Integer
    Dim Y As Integer
    sqlADP = "select * from tbAvaliacaoDesempenho where codcoligada = '" & vCodcoligada & "' order by id"
    rsADP.Open sqlADP, cnBanco, adOpenKeyset, adLockReadOnly
    If rsADP.RecordCount = 0 Then Exit Function
    For X = 0 To rsADP.RecordCount - 1
        vADP(X, 0) = rsADP.Fields(1)
        vADP(X, 1) = rsADP.Fields(2)
        rsADP.MoveNext
    Next
    rsADP.Close
    Set rsADP = Nothing
    montaDadosADP
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Public Function montaDadosADP()
On Error GoTo Err
    Dim rsMontaDadosADP As New ADODB.Recordset
    Dim SqlMontaDadosADP As String
    
    Dim rsVerificaADP As New ADODB.Recordset
    Dim SqlVerificaADP As String
    Dim diasProximaADP As Integer
    
    'Todos os colaboradors com a quantidade de dias que estão na matriz
    SqlMontaDadosADP = "select a.id, a.nomecolaborador, b.codmatriz, b.data, DATEDIFF(DAY,b.data,GETDATE()) from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf where a.ativo = 'S' and b.ativo = 'S'"
    rsMontaDadosADP.Open SqlMontaDadosADP, cnBanco, adOpenKeyset, adLockReadOnly
    For X = 1 To rsMontaDadosADP.RecordCount
        SqlVerificaADP = "Select * from tblistaADP where codcoligada = '" & vCodcoligada & "' and codcolaborador = '" & rsMontaDadosADP.Fields(0) & "' and statusavaliacao is null or codcoligada = '" & vCodcoligada & "' and codcolaborador = '" & rsMontaDadosADP.Fields(0) & "' and statusavaliacao <> 'Concluido'"
        rsVerificaADP.Open SqlVerificaADP, cnBanco, adOpenKeyset, adLockOptimistic
        'SE FOR = 0 NAO EXISTE AVALIACAO EM ABERTO PARA O COLABORADOR
        'ENTRA NA CONDIÇÃO ABAIXO
        If rsVerificaADP.RecordCount = 0 Then
            diasTrabalhados = rsMontaDadosADP.Fields(4)
            avaliarAKDA = achaDias(rsMontaDadosADP.Fields(0))
            If Val(diasTrabalhados) > Val(avaliarAKDA) Then
                diasProximaADP = Val(diasTrabalhados / avaliarAKDA) * avaliarAKDA + avaliarAKDA
            Else
                diasProximaADP = avaliarAKDA
            End If
            ' AKI CHAMA ROTINA DE GRAVAÇÃO
            rsVerificaADP.AddNew
            rsVerificaADP.Fields(1) = rsMontaDadosADP.Fields(0)
            rsVerificaADP.Fields(2) = tipoADP
            rsVerificaADP.Fields(3) = avaliarAKDA
            'Teste para corrigir o erro de 1 dia na avaliação de desempenho
            rsVerificaADP.Fields(5) = rsMontaDadosADP.Fields(3) + (diasProximaADP - 1) 'Teste para corrigir o erro de 1 dia na avaliação de desempenho
            rsVerificaADP.Fields(6) = rsMontaDadosADP.Fields(3) + (diasProximaADP - 3)
            rsVerificaADP.Fields(23) = "-"
            rsVerificaADP.Fields(24) = "S"
            rsVerificaADP.Fields(26) = vCodcoligada 'Codigo da coligada
            rsVerificaADP.Update
            '-------------------------------
        End If
        rsVerificaADP.Close
        rsMontaDadosADP.MoveNext
    Next
    rsMontaDadosADP.Close
    Set rsMontaDadosADP = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Public Function achaDias(vCodColab As String)
On Error GoTo Err
    Dim rsAchaDias As New ADODB.Recordset
    Dim SqlAchaDias As String
    Dim X As Integer
    
    achaDias = 0
    
    SqlAchaDias = "select a.id,a.codcolaborador,a.dias,a.statusavaliacao from tbListaADP as a where a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador = '" & vCodColab & "' and statusavaliacao = 'concluido' order by a.dias desc"
    rsAchaDias.Open SqlAchaDias, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsAchaDias.EOF Then achaDias = rsAchaDias.Fields(2)
    rsAchaDias.Close
    Set rsAchaDias = Nothing
    '--> SE ENCONTRAR AVALIAÇÕES JA CONCLUIDAS NO SISTEMA
    If achaDias > 0 Then
        For X = 0 To 10
            If vADP(X, 0) = "" Then Exit Function
            If Val(vADP(X, 0)) > achaDias Then
                achaDias = vADP(X, 0)
                tipoADP = vADP(X, 1)
                If diasTrabalhados < achaDias Then Exit Function
            End If
        Next
    '--> SE NÃO ENCONTRAR AVALIAÇÕES CONCLUIDAS NO SISTEMA
    Else
        For X = 0 To 10
            If vADP(X, 0) = "" Then Exit Function
            achaDias = vADP(X, 0)
            tipoADP = vADP(X, 1)
            If diasTrabalhados < achaDias Then Exit Function
        Next
    End If
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function
