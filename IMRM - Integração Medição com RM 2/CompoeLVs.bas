Attribute VB_Name = "CompoeLVs"
Public apontaLV As Integer
Public indiceVarGlobal As Integer 'quantas colunas vai ter a variavel global
Public checaFiltro As Boolean
Public vADP(10, 1) As String
Public diasTrabalhados As Integer
Public avaliarAKDA As Integer, vSegment As Integer
Public tipoADP As String
Public vQualColunaStatusMedicao As Integer

Public Sub MontaLV(QualLV As Integer)
        'On Error GoTo TrataErro
        If vAvisos = "" Then
            Msgbox "Local de Estoque n�o ativo. Acesse: Configura��es|Sistema|Parametriza��es|Gerais e informe", vbCritical, "Ferramentaria"
            Exit Sub
        ElseIf vBancoSAP = "" Then
            Msgbox "Par�metros de integra��o n�o informados. Acesse: Configura��es|Sistema|Parametriza��es|Integra��o e informe", vbCritical, "IMRM"
            Exit Sub
        ElseIf vCodcoligada = 0 Then
            Msgbox "Coligada n�o cadastrada. Acesse: Configura��es|Sistema|Coligadas e informe", vbCritical, "IMRM"
            Exit Sub
        End If
        
        'MEDI��O DE TERCEIROS
        If QualLV = 0 Then
            If vCodVenRM = "" Then
                Msgbox "Usu�rio n�o vinculado ao TOTVS RM. Acesse: Configura��es|usu�rios e vincule", vbCritical, "IMRM"
                Exit Sub
            End If
            
            'Set chamaForm = New frmEmprestimo
            Formulario = "Medi��o de Terceiro"
            LegendaExc = "Medi��o de Terceiro" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then
                    Unload MeuLV
                End If
                carregaTABS "ID_APROP_MEDICAOTERCEIRO", "ID_FUNC", "ID_FUNC", "FCFO", "ID_FUNC", "PSECAO", "ID_APROP_PERIODO", "ID_PRJ_PROJETO", "ID_APROP_APROVACAO", "ID_APROP_APROVACAOSTATUS", "tbMedicoesTerceiro", "TMOVRELAC", "TMOV", "", ""
                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.CODIGO,d.NOME AS EMPRESA,(e.NOME + EXECUTANTEANONIMO) AS EXECUTANTE,f.DESCRICAO + ' - ' + f.CGC AS SECAO,CONVERT(VARCHAR,g.DTINICIAL,103) + ' a ' + CONVERT(VARCHAR,g.DTFINAL,103) AS PERIODO,h.CODIGO AS PROJETO,a.COMPETENCIA,a.VALORTOTAL,CONVERT(VARCHAR(10),a.DTCADASTRO,103) AS DATACADASTRO,a.NOTAFISCAL,j.NOME AS STATUS,d.CGCCFO AS CNPJ,cast(substring(a.codsecao,1,2) as int) as codfilial, a.codcfo,k.status,M.NUMEROMOV, M.CODTMV " & _
                                 "from " & vBancoSAP & ".DBO.ID_APROP_MEDICAOTERCEIRO as a WITH (NOLOCK) LEFT JOIN " & vBancoSAP & ".DBO.ID_FUNC as b WITH (NOLOCK) ON a.IDRESPONSAVEL = b.IDINFO LEFT JOIN " & vBancoSAP & ".DBO.ID_FUNC as c WITH (NOLOCK) ON a.IDGERENTE = c.IDINFO LEFT JOIN " & vBancoSAP & ".DBO.FCFO as d WITH (NOLOCK) ON a.CODCFO = d.CODCFO LEFT JOIN " & vBancoSAP & ".DBO.ID_FUNC as e WITH (NOLOCK) ON e.IDINFO = a.IDEXECUTANTE LEFT JOIN " & vBancoSAP & ".DBO.PSECAO as f WITH (NOLOCK) ON f.CODIGO = a.CODSECAO " & _
                                 "LEFT JOIN " & vBancoSAP & ".DBO.ID_APROP_PERIODO as g WITH (NOLOCK) ON g.ID = a.IDPERIODO LEFT JOIN " & vBancoSAP & ".DBO.ID_PRJ_PROJETO as h WITH (NOLOCK) ON h.ID = a.IDPROJETO LEFT JOIN " & vBancoSAP & ".DBO.ID_APROP_APROVACAO as i WITH (NOLOCK) ON i.IDMEDICAOTERCEIRO = a.ID LEFT JOIN " & vBancoSAP & ".DBO.ID_APROP_APROVACAOSTATUS as j WITH (NOLOCK) ON i.IDSTATUS = j.ID LEFT JOIN tbMedicoesTerceiro as k WITH (NOLOCK) on a.CODIGO = k.codigo COLLATE SQL_Latin1_General_CP1_CI_AS " & _
                                 "LEFT JOIN " & vBancoSAP & ".DBO.TMOVRELAC AS L ON K.idmovintegracao = L.IDMOVORIGEM LEFT JOIN " & vBancoSAP & ".DBO.TMOV AS M ON L.IDMOVDESTINO = M.IDMOV order by a.id desc"
'                If FiltroGeral = "Terceiros" Then SqlLV = "Select top " & LimiteLinhas & " a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,dataemprestimo,a.idmov,a.numeromov,a.serie,a.nomequememprestou,a.codusuariorm,a.status,c.CODSITUACAO,'' from tbEmprestimo as a left join " & vBancoSAP & ".dbo.PFUNC as c on a.chapa = c.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS where a.status = 'E' and a.localestoque = " & Val(vLocalEstoque) & "  order by a.dataemprestimo desc,a.idmov desc"
'                If FiltroGeral = "Devolu��es" Then SqlLV = "Select top " & LimiteLinhas & " a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,dataemprestimo,a.idmov,a.numeromov,a.serie,a.nomequememprestou,a.codusuariorm,a.status,c.CODSITUACAO from tbEmprestimo as a left join " & vBancoSAP & ".dbo.PFUNC as c on a.chapa = c.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS where a.status <> 'E' and a.localestoque = " & Val(vLocalEstoque) & "  order by a.dataemprestimo desc,a.idmov desc"
'                If FiltroGeral = "Devolu��es" Then SqlLV = "Select top " & LimiteLinhas & " a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,B.datadevolucao,b.idmovemp,a.numeromov,a.serie,B.nomequememprestou,a.codusuariorm,'D',c.CODSITUACAO,a.idmov from tbDevolucao as a left join CORPORERM_OFF.dbo.PFUNC as c on a.chapa = c.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS " & _
'                                 "INNER JOIN tbDevolucaoItens AS B ON A.idmov = B.idmov where a.localestoque =  " & Val(vLocalEstoque) & "  group by a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,B.datadevolucao,a.idmov,a.numeromov,a.serie,B.nomequememprestou,a.codusuariorm,c.CODSITUACAO,b.idmovemp order by b.datadevolucao desc,a.idmov desc"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.ListView1.CheckBoxes = True
            MeuLV.cmdconsulta(4).Visible = True
            MeuLV.cmdconsulta(5).Visible = True
            MeuLV.cmdconsulta(6).Visible = True
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "C�d Medi��o", "Empresa", "Executante", "Se��o", "Per�odo", "Projeto", "Compet�ncia", "Valor Total", "D.Cad.", "NF", "Status", "CNPJ", "Filial", "ID Fornec", "Status Env.", "NF Totvs", "Tipo Mov.", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Medi��o de Terceiro"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 7, "N", "N", "", "N", "N", "S", "D"
                PersonaColLV 14, "N", "N", "", "S", "N", "N", "C"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(20).Picture
            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(14).Picture
            Set MeuLV.cmdconsulta(6).PictureNormal = MeuLV.ImageList1.ListImages(22).Picture
            MeuLV.cmdconsulta(4).ToolTipText = "Marcar todos"
            MeuLV.cmdconsulta(5).ToolTipText = "Exportar para o RM"
            MeuLV.cmdconsulta(6).ToolTipText = "Bloqueia Medi��o"
            Exit Sub
        'MEDI��O DE PJ/MENSAL
        ElseIf QualLV = 1 Then
            If vCodVenRM = "" Then
                Msgbox "Usu�rio n�o vinculado ao TOTVS RM. Acesse: Configura��es|usu�rios e vincule", vbCritical, "IMRM"
                Exit Sub
            End If
            
            'Set chamaForm = New frmEmprestimo
            Formulario = "Medi��o PJ/Mensal"
            LegendaExc = "Medi��o PJ/Mensal" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then
                    Unload MeuLV
                End If
                carregaTABS "ID_APROP_APROVACAO", "ID_APROP_PERIODO", "ID_FUNC", "ID_FUNC", "ID_FUNC", "ID_APROP_APROVACAOSTATUS", "ID_APROP_MEDICAOTIPO", "ID_APROP_MEDICAO", "PEXTERNO", "FCFO", "tbMedicoesPJ", "TMOVRELAC", "TMOV", "", ""
                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " REPLICATE('0', 5 - LEN(h.ID)) + RTrim(h.ID) AS MEDICAO,j.CGCCFO AS CNPJ,CASE a.IDINFO WHEN 0 THEN (SELECT EXECUTANTEANONIMO FROM  " & vBancoSAP & ".DBO.ID_APROP_MEDICAOTERCEIRO WHERE ID_APROP_MEDICAOTERCEIRO.ID = a.IDMEDICAOTERCEIRO ) ELSE c.NOME END AS COLABORADOR,d.NOME AS APROVADORPRIMARIO,CONVERT(VARCHAR(10), a.DTAPROVADORPRIMARIO, 103) AS DATAAPROVADORPRIMARIO,e.NOME AS APROVADORSECUNDARIO,CONVERT(VARCHAR(10), a.DTAPROVADORSECUNDARIO, 103) AS DATAAPROVADORSECUNDARIO,CONVERT(VARCHAR,b.DTINICIAL,103) + ' a ' + CONVERT(VARCHAR,b.DTFINAL,103) AS PERIODO,f.NOME AS STATUS,g.NOME AS MEDICAOTIPO,a.HORAS,a.VLRHORA,h.TOTAL,a.NOTAFISCAL,h.codfilial,h.codcfo,k.status,a.idinfo,j.nome as empresa,M.NUMEROMOV, M.CODTMV, " & _
                                                      "CASE WHEN A.DTAPROVADORSECUNDARIO IS NULL  then CASE WHEN DatePart(day, A.DTAPROVADORPRIMARIO) <= 10 Then '01/' + convert(varchar(10),DatePart(month, getdate()) - 1,103)  + '/' + convert(varchar(10),DatePart(year, getdate()),103) WHEN DatePart(day, A.DTAPROVADORPRIMARIO) > 10 Then CASE WHEN DatePart(month, A.DTAPROVADORPRIMARIO) < DatePart(month, getdate()) Then '01/' + convert(varchar(10),DatePart(month, getdate()) - 1,103)  + '/' + convert(varchar(10),DatePart(year, getdate()),103) ELSE CASE WHEN DatePart(year, A.DTAPROVADORPRIMARIO) < DatePart(year, getdate()) Then " & _
                                                      "'01/' + convert(varchar(10),DatePart(month, getdate()) - 1,103)  + '/' + convert(varchar(10),DatePart(year, getdate()),103) ELSE '01/' + convert(varchar(10),DatePart(month, getdate()),103) + '/' + convert(varchar(10),DatePart(year, getdate()),103) END END END ELSE CASE WHEN DatePart(day, A.DTAPROVADORSECUNDARIO) <= 10 Then '01/' + convert(varchar(10),DatePart(month, getdate()) - 1,103)  + '/' + convert(varchar(10),DatePart(year, getdate()),103) WHEN DatePart(day, A.DTAPROVADORSECUNDARIO) > 10 Then CASE WHEN DatePart(month, A.DTAPROVADORSECUNDARIO) < DatePart(month, getdate()) Then '01/' + convert(varchar(10),DatePart(month, getdate()) - 1,103)  + '/' + convert(varchar(10),DatePart(year, getdate()),103) " & _
                                                      "ELSE CASE WHEN DatePart(year, A.DTAPROVADORSECUNDARIO) < DatePart(year, getdate()) Then '01/' + convert(varchar(10),DatePart(month, getdate()) - 1,103)  + '/' + convert(varchar(10),DatePart(year, getdate()),103) ELSE '01/' + convert(varchar(10),DatePart(month, getdate()),103) + '/' + convert(varchar(10),DatePart(year, getdate()),103) END END END END COMPETENCIA " & _
                                                      "from " & vBancoSAP & ".DBO.ID_APROP_APROVACAO a WITH (NOLOCK) INNER JOIN " & vBancoSAP & ".DBO.ID_APROP_PERIODO b WITH (NOLOCK) ON a.IDPERIODO = b.ID and a.MEDICAOAVULSA != 3 AND a.CONDICAO = 'PJ' LEFT JOIN " & vBancoSAP & ".DBO.ID_FUNC c WITH (NOLOCK) ON a.IDINFO = c.IDINFO INNER JOIN " & vBancoSAP & ".DBO.ID_FUNC d WITH (NOLOCK) ON a.IDAPROVADORPRIMARIO = d.IDINFO LEFT JOIN " & vBancoSAP & ".DBO.ID_FUNC e WITH (NOLOCK) ON a.IDAPROVADORSECUNDARIO = e.IDINFO INNER JOIN " & vBancoSAP & ".DBO.ID_APROP_APROVACAOSTATUS f WITH (NOLOCK) ON a.IDSTATUS = f.ID INNER JOIN " & vBancoSAP & ".DBO.ID_APROP_MEDICAOTIPO g WITH (NOLOCK) ON a.MEDICAOAVULSA = g.ID " & _
                                                      "left join " & vBancoSAP & ".DBO.ID_APROP_MEDICAO h WITH (NOLOCK) on a.ID = h.IDAPROVACAO left join " & vBancoSAP & ".DBO.PEXTERNO as i WITH (NOLOCK) on c.IDINFO = i.CODEXTERNO LEFT JOIN  " & vBancoSAP & ".DBO.FCFO as j WITH (NOLOCK) on i.CODCFO = j.CODCFO LEFT JOIN tbMedicoesPJ as k WITH (NOLOCK) on h.ID = k.codigo LEFT JOIN " & vBancoSAP & ".DBO.TMOVRELAC AS L ON K.idmovintegracao = L.IDMOVORIGEM LEFT JOIN " & vBancoSAP & ".DBO.TMOV AS M ON L.IDMOVDESTINO = M.IDMOV order by h.id desc"
'                If FiltroGeral = "Terceiros" Then SqlLV = "Select top " & LimiteLinhas & " a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,dataemprestimo,a.idmov,a.numeromov,a.serie,a.nomequememprestou,a.codusuariorm,a.status,c.CODSITUACAO,'' from tbEmprestimo as a left join " & vBancoSAP & ".dbo.PFUNC as c on a.chapa = c.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS where a.status = 'E' and a.localestoque = " & Val(vLocalEstoque) & "  order by a.dataemprestimo desc,a.idmov desc"
'                If FiltroGeral = "Devolu��es" Then SqlLV = "Select top " & LimiteLinhas & " a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,dataemprestimo,a.idmov,a.numeromov,a.serie,a.nomequememprestou,a.codusuariorm,a.status,c.CODSITUACAO from tbEmprestimo as a left join " & vBancoSAP & ".dbo.PFUNC as c on a.chapa = c.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS where a.status <> 'E' and a.localestoque = " & Val(vLocalEstoque) & "  order by a.dataemprestimo desc,a.idmov desc"
'                If FiltroGeral = "Devolu��es" Then SqlLV = "Select top " & LimiteLinhas & " a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,B.datadevolucao,b.idmovemp,a.numeromov,a.serie,B.nomequememprestou,a.codusuariorm,'D',c.CODSITUACAO,a.idmov from tbDevolucao as a left join CORPORERM_OFF.dbo.PFUNC as c on a.chapa = c.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS " & _
'                                 "INNER JOIN tbDevolucaoItens AS B ON A.idmov = B.idmov where a.localestoque =  " & Val(vLocalEstoque) & "  group by a.chapa,a.nome,a.codfuncao,a.nomefuncao,a.codsecao,a.nomesecao,B.datadevolucao,a.idmov,a.numeromov,a.serie,B.nomequememprestou,a.codusuariorm,c.CODSITUACAO,b.idmovemp order by b.datadevolucao desc,a.idmov desc"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.ListView1.CheckBoxes = True
            MeuLV.cmdconsulta(4).Visible = True
            MeuLV.cmdconsulta(5).Visible = True
            MeuLV.cmdconsulta(6).Visible = True
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "Medi��o", "CNPJ", "Colaborador", "Ap. Prim.", "DT. Ap. Prim.", "Ap. Secun.", "DT. Ap. Secun.", "Per�odo", "Status", "Tipo Medi��o", "Horas", "Vlr. Hora", "Total", "NF", "Filial", "ID Fornec", "Status Env.", "IDINFO", "Empresa", "NF Totvs", "Tipo Mov.", "Compet�ncia", "", "", "", ""
            DimensionaLV " PJ/Mensal"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 11, "N", "N", "", "N", "N", "S", "D"
                PersonaColLV 12, "N", "N", "", "N", "N", "S", "D"
                PersonaColLV 16, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(20).Picture
            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(14).Picture
            Set MeuLV.cmdconsulta(6).PictureNormal = MeuLV.ImageList1.ListImages(22).Picture
            MeuLV.cmdconsulta(4).ToolTipText = "Marcar todos"
            MeuLV.cmdconsulta(5).ToolTipText = "Exportar para o RM"
            MeuLV.cmdconsulta(6).ToolTipText = "Bloqueia Medi��o"
            Exit Sub
        'GRUPO DE CRIT�RIOS DE AVALIA��O DE FORNECEDORES
        ElseIf QualLV = 2 Then
            'Set chamaForm = New frmGrupoAvFornec
            Formulario = "Grupo de Crit�rios de Avalia��o de Fornecedores"
            LegendaExc = "Grupo de Crit�rios para Credenciamento de Fornecedores" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then
                    Unload MeuLV
                End If
                If FiltroGeral = "Todos" Then SqlLV = "Select a.idavfornecGrup,a.nomeavfornecgrup,a.ativo from tbAvFornecGrup as a"
                If FiltroGeral = "Ativos" Then SqlLV = "Select a.idavfornecGrup,a.nomeavfornecgrup,a.ativo from tbAvFornecGrup as a Where a.ativo = 'S'"
                If FiltroGeral = "N�o ativos" Then SqlLV = "Select a.idavfornecGrup,a.nomeavfornecgrup,a.ativo from tbAvFornecGrup as a where a.ativo <> 'S'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "ID", "Grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Grupo de Crit�rios para Credenciamento de Fornecedores"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Exit Sub
        'CRIT�RIOS DE AVALIA��O DE FORNECIMENTO
        ElseIf QualLV = 3 Then
            'Set chamaForm = New frmCriterioFornec
            Formulario = "Crit�rios de Avalia��o de Fornecimento"
            LegendaExc = "Crit�rios para Credenciamento de Fornecedores" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "Select a.idcriterioavfornec,a.nomecriterioavfornec,a.criticidade,a.ativo from tbCriterioAvFornec as a"
                If FiltroGeral = "Ativos" Then SqlLV = "Select a.idcriterioavfornec,a.nomecriterioavfornec,a.criticidade,a.ativo from tbCriterioAvFornec as a where a.ativo = 'S'"
                If FiltroGeral = "N�o ativos" Then SqlLV = "Select a.idcriterioavfornec,a.nomecriterioavfornec,a.criticidade,a.ativo from tbCriterioAvFornec as a where a.ativo <> 'S'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "ID", "Nome", "Criticidade", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Crit�rios para Credenciamento de Fornecedores"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Exit Sub
        'FORNECEDORES
        ElseIf QualLV = 4 Then
            'Set chamaForm = New frmAvFornecedor
            Formulario = "Fornecedores"
            LegendaExc = "Fornecedores" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "Select top " & LimiteLinhas & " a.CardCode as ID,CASE WHEN b.ativo IS NULL or B.ATIVO = 'N' THEN 'N' ELSE 'S' END ATIVO,a.CardName as Nome,a.Address + ', ' + a.StreetNo as Endereco,a.ZipCode,a.City as Cidade,a.Block as Bairro,a.State1 as UF,a.E_Mail,CASE WHEN b.situacao is null then 'N' Else 'S' end as SITUACAO," & _
                                                      "CASE WHEN b.status is null then '-' ELSE B.status END STATUS,CASE WHEN B.grupo IS NULL THEN '-' ELSE B.grupo END GRUPO,b.nomeavfornecgrup,CONVERT (VARCHAR, b.datacredenciamento, 103) as datacredenciamento from " & vBancoSAP & ".DBO.OCRD as a LEFT JOIN tbFornecedores " & _
                                                      "AS B ON A.CardCode = B.idfornecedor COLLATE SQL_Latin1_General_CP1_CI_AS where a.validFor = 'Y' and a.CardType = 'S'"
                If FiltroGeral = "Ativos" Then SqlLV = "Select top " & LimiteLinhas & " a.CardCode as ID,CASE WHEN b.ativo IS NULL or B.ATIVO = 'N' THEN 'N' ELSE 'S' END ATIVO,a.CardName as Nome,a.Address + ', ' + a.StreetNo as Endereco,a.ZipCode,a.City as Cidade,a.Block as Bairro,a.State1 as UF,a.E_Mail,CASE WHEN b.situacao is null then 'N' Else 'S' end as SITUACAO," & _
                                                      "CASE WHEN b.status is null then '-' ELSE B.status END STATUS,CASE WHEN B.grupo IS NULL THEN '-' ELSE B.grupo END GRUPO,b.nomeavfornecgrup,CONVERT (VARCHAR, b.datacredenciamento, 103) as datacredenciamento from " & vBancoSAP & ".DBO.OCRD as a LEFT JOIN tbFornecedores " & _
                                                      "AS B ON A.CardCode = B.idfornecedor COLLATE SQL_Latin1_General_CP1_CI_AS where a.validFor = 'Y' and a.CardType = 'S' and B.ATIVO = 'S'"
                If FiltroGeral = "N�o ativos" Then SqlLV = "Select top " & LimiteLinhas & " a.CardCode as ID,CASE WHEN b.ativo IS NULL or B.ATIVO = 'N' THEN 'N' ELSE 'S' END ATIVO,a.CardName as Nome,a.Address + ', ' + a.StreetNo as Endereco,a.ZipCode,a.City as Cidade,a.Block as Bairro,a.State1 as UF,a.E_Mail,CASE WHEN b.situacao is null then 'N' Else 'S' end as SITUACAO," & _
                                                      "CASE WHEN b.status is null then '-' ELSE B.status END STATUS,CASE WHEN B.grupo IS NULL THEN '-' ELSE B.grupo END GRUPO,b.nomeavfornecgrup,CONVERT (VARCHAR, b.datacredenciamento, 103) as datacredenciamento from " & vBancoSAP & ".DBO.OCRD as a LEFT JOIN tbFornecedores " & _
                                                      "AS B ON A.CardCode = B.idfornecedor COLLATE SQL_Latin1_General_CP1_CI_AS where a.validFor = 'Y' and a.CardType = 'S' and b.ativo is null or a.validFor = 'Y' and a.CardType = 'S' and b.ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.ListView1.CheckBoxes = True
            MeuLV.cmdconsulta(4).Visible = False
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = True
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "ID", "Ativo", "Nome", "Endere�o", "CEP", "Cidade", "Bairro", "UF", "Email", "Credenciado?", "Situa��o", "Grupo (Credenciamento)", "Grupo (Recebimento)", "Data Credenciamento", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Ativa��o/Credenciamento de Fornecedores"
            MontaCabecalhoLV
            MontaDadosLV "S" 'Coloca zeros a esquerda na primeira coluna
            If checaFiltro = True Then
                PersonaColLV 1, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 9, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 10, "N", "S", "", "N", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            MeuLV.cmdconsulta(5).ToolTipText = "Avaliar Fornecedor"
            
            Set MeuLV.cmdconsulta(11).PictureNormal = MeuLV.ImageList1.ListImages(19).Picture
            MeuLV.cmdconsulta(11).ToolTipText = "Selecionar Grupo de crit�rios de Avalia��o de Fornecimento"
            
            Exit Sub
        'RECEBIMENTO NF
        ElseIf QualLV = 5 Then
            'Set chamaForm = New frmRecebePedCompra
            Formulario = "Recebimento de Ordem de Compra"
            LegendaExc = "Recebimento de Pedido de Compra" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.docnum as Ordem_compra,CASE WHEN a.DocStatus = 'C' THEN 'Fechado' else 'Aberto' end Status_SAP,CONVERT (VARCHAR, a.DocDate, 103) as Data_documento,a.CardCode as ID_Fornecedor,a.CardName as nome_fornecedor,CASE WHEN b.statusoc IS NULL or b.statusoc = 'N' THEN 'N' WHEN b.statusoc='7' THEN '7' else 'S' end Statusoc,b.notaOC,b.dataavoc," & _
                                                      "CASE WHEN c.idavfornecgrup IS NULL  THEN 'N' else 'S' end grupoAV,b.avaliadopor,a.segment,d.idclassificacao from " & vBancoSAP & ".DBO.OPOR as a LEFT JOIN tbOCStatus as b on a.DocNum = b.docnum and a.Segment = b.segment inner join tbfornecedores as c on a.CardCode COLLATE SQL_Latin1_General_CP1_CI_AS = c.idfornecedor left join tbClassificacao as d on b.notaoc >= d.de and b.notaoc <= d.para where a.DocDate >= '" & vInicioAvOC & "' and a.CANCELED = 'N' and c.status = 'Credenciado' and c.ativo = 'S' order by a.DocDate desc"
'                If FiltroGeral = "Ativos" Then SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato where b.ativo = 'S'"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "select b.codfo,a.nome,b.pedido,c.nome,c.telefone,b.descricao,b.datafo,b.datadevcp,b.proposta,b.quantidade,b.valorunit,(b.quantidade*b.valorunit) as valortotal,b.pedido,b.fce,b.statusfo,b.ativo from tbclifor as a inner join tbfo  as b on a.codclifor=b.codclifor left join tbcontatos as c on b.codclifor = c.codclifor and b.codcontato = c.codcontato where b.ativo <> 'S'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(4).Visible = False
            MeuLV.cmdconsulta(6).Visible = False
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "Ordem Compra", "Status OC SAP", "Data Emiss�o OC", "Id Fornecedor", "Nome Fornecedor", "Status da Avalia��o", "%", "Data Avalia��o", "Grupo Avalia��o", "Avaliado Por", "Segmento", "Classifica��o", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Recebimento de Ordem de Compra"
            MontaCabecalhoLV
            MontaDadosLV "N" 'Coloca zeros a esquerda na primeira coluna
            If checaFiltro = True Then
                PersonaColLV 5, "N", "S", "", "S", "N", "N", "E"
                PersonaColLV 6, "S", "S", "%", "N", "N", "S", "D"
                PersonaColLV 8, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 11, "N", "S", "", "N", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
'            Set MeuLV.cmdconsulta(9).PictureNormal = MeuLV.ImageList1.ListImages(7).Picture
'            MeuLV.cmdconsulta(9).ToolTipText = "Receber FO"
'            Set MeuLV.cmdconsulta(11).PictureNormal = MeuLV.ImageList1.ListImages(22).Picture
'            MeuLV.cmdconsulta(11).ToolTipText = "Editar FCE"
            Exit Sub
'        'Notas das Avalia��es dos Fornecedores
        ElseIf QualLV = 6 Then
'            Set chamaForm = New frmFCECons
            Formulario = "Notas das Avalia��es dos Fornecedores"
            LegendaExc = "Notas das Avalia��es dos Fornecedores" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select * from TempNotaFornec"
'                If FiltroGeral = "Ativos" Then SqlLV = "select b.dataabertura,b.fce,c.nome[cliente],d.nome[contato],d.telefone,b.dataentrega,b.pintura,b.transporte,b.materiaprima,b.fabricacao,b.reparo,a.ativo from tbfo as a inner join tbfce as b on b.fce = a.fce left join tbclifor as c on a.codclifor=c.codclifor left join tbcontatos as d on c.codclifor = d.codclifor and d.codcontato = a.codcontato where a.ativo = 'S' order by b.fce"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "select b.dataabertura,b.fce,c.nome[cliente],d.nome[contato],d.telefone,b.dataentrega,b.pintura,b.transporte,b.materiaprima,b.fabricacao,b.reparo,a.ativo from tbfo as a inner join tbfce as b on b.fce = a.fce left join tbclifor as c on a.codclifor=c.codclifor left join tbcontatos as d on c.codclifor = d.codclifor and d.codcontato = a.codcontato where a.ativo <> 'S' order by b.fce"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(4).Visible = False
            MeuLV.cmdconsulta(5).Visible = False
            MeuLV.cmdconsulta(6).Visible = False
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "ID Fornecedor", "Nome Fornecedor", "Nota Geral %", "Classifica��o", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Notas das Avalia��es dos Fornecedores"
            MontaCabecalhoLV
            MontaDadosLV "N"
            If checaFiltro = True Then
                PersonaColLV 2, "S", "N", "%", "N", "N", "S", "D"
                PersonaColLV 3, "S", "S", "", "N", "N", "S", "D"
                'PersonaColLV 11, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(8).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Nova LM - Lista de Materiais"
'            MeuLV.cmdconsulta(5).ToolTipText = "Consultar FCE - Ficha de Controle de Encomenda"
            Exit Sub
        'CADASTRO DE DESENHOS
        ElseIf QualLV = 7 Then
'            Set chamaForm = New frmDesenhos
'            Formulario = "Cadastro de Desenhos"
'            LegendaExc = "Cadastro de Desenhos" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "Select top " & LimiteLinhas & " a.iddesenho,a.desenho,a.revisao,b.fce,b.projeto,a.datacadastro,a.tipo,a.ativo from  tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto where a.codcoligada = '" & vCodcoligada & "' order by b.fce desc,b.projeto"
'                If FiltroGeral = "Ativos" Then SqlLV = "Select top " & LimiteLinhas & " a.iddesenho,a.desenho,a.revisao,b.fce,b.projeto,a.datacadastro,a.tipo,a.ativo from  tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto where a.codcoligada = '" & vCodcoligada & "' and a.ativo='S' order by b.fce desc,b.projeto"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "Select top " & LimiteLinhas & " a.iddesenho,a.desenho,a.revisao,b.fce,b.projeto,a.datacadastro,a.tipo,a.ativo from  tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto where a.codcoligada = '" & vCodcoligada & "' and a.ativo='N' order by b.fce desc,b.projeto"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Identificador", "Desenho", "Rev.", "FCE", "Projeto", "Data Cadastro", "Tipo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Cadastro de Desenhos"
'            MontaCabecalhoLV
'            MontaDadosLV "S" ' Zero a direita na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "N", "", "N", "N", "N", "E"
'                PersonaColLV 7, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.cmdconsulta(6).ToolTipText = "Cancelar treinamento"
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Exit Sub
        'LM - LISTA DE MATERIAIS
        ElseIf QualLV = 8 Then
'            Set chamaForm = New frmLM
'            Formulario = "LM's - Listas de Materiais"
'            LegendaExc = "LM - Lista de Material" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 2 'Com quantas colunas que a varglobal ir� trabalhar
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select right('0000' + rtrim(fce),4),codlm,dataabertura,descricao,ativo from tblm order by tblm.fce, tblm.codlm"
'                If FiltroGeral = "Ativos" Then SqlLV = "select right('0000' + rtrim(fce),4),codlm,dataabertura,descricao,ativo from tblm where tblm.ativo = 'S' order by tblm.fce, tblm.codlm"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "select right('0000' + rtrim(fce),4),codlm,dataabertura,descricao,ativo from tblm where tblm.ativo <> 'S' by tblm.fce, tblm.codlm"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "FCE", "LM", "Data Abertura", "Descri��o", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "LM's - Listas de Materiais"
'            MontaCabecalhoLV
'            MontaDadosLV "N"
'            If checaFiltro = True Then
'                PersonaColLV 1, "S", "S", "", "N", "S", "N", "D"
'                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            MeuLV.cmdconsulta(5).ToolTipText = "Editar LM - Lista de Materiais"
'            Exit Sub
        'MP - M�todos e Processos
        ElseIf QualLV = 9 Then
'            Set chamaForm = New frmMPCompleto
'            Formulario = "M�todo & Processo"
'            LegendaExc = "M�todo & Processo" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1 'Com quantas colunas que a varglobal ir� trabalhar
'
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produ��o' when a.status = 3 then 'Expedi��o' else 'Planejamento' end as status from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto " & _
'                                                      "left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e " & _
'                                                      "on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status order by a.idprogramacao desc"
'                If FiltroGeral = "Planejamento" Then SqlLV = "select a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produ��o' when a.status = 3 then 'Expedi��o' else 'Planejamento' end as status from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto " & _
'                                                       "left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e " & _
'                                                       "on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos  and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao where a.status = 1 group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status order by a.idprogramacao"
'                If FiltroGeral = "Produ��o" Then SqlLV = "select a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produ��o' when a.status = 3 then 'Expedi��o' else 'Planejamento' end as status from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto " & _
'                                                       "left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e " & _
'                                                       "on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos  and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao where a.status > 1 and a.status < 3 group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status order by a.idprogramacao"
'                If FiltroGeral = "Expedi��o" Then SqlLV = "select a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,min(e.desenho) as Desenho,a.ativo,max(g.idretrabalho) as retrabalho,case when a.status = 1 then 'Planejamento' when a.status > 1 and a.status < 3 then 'Produ��o' when a.status = 3 then 'Expedi��o' else 'Planejamento' end as status from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto " & _
'                                                           "left join tbMPItens as c on a.idprogramacao = c.idprogramacao left join tbitemlm as d on SUBSTRING(c.desenhos,1,2) = d.codlm and replace(SUBSTRING(c.desenhos,3,4),';','') = d.codseq and b.fce = d.fce left join tbDesenhos as e " & _
'                                                           "on d.codigodes = e.iddesenho left join tbos as f on c.idos = f.idos  and c.revisaoos = f.revisao left join tbretrabalho as g on a.idprogramacao = g.idprogramacao where a.status = 3 group by a.idprogramacao,c.idos,f.revisao,a.dataprogramacao,b.fce,b.projeto,a.responsavel,a.ativo,a.status order by a.idprogramacao"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(0).Visible = True
'            MeuLV.cmdconsulta(9).Visible = True
'            MeuLV.cmdconsulta(11).Visible = True
'            MeuLV.cmdconsulta(12).Visible = True
'            QtdColReal = 0
'            MontaCabLV "Planejamento", "OS n�", "Rev.", "Data", "FCE", "Projeto", "Respons�vel", "Desenho", "Ativo", "Retrabalho", "Status", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "M�todos e Processos"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'                PersonaColLV 1, "N", "N", "", "N", "S", "N", "E"
'                PersonaColLV 8, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 9, "S", "N", "", "N", "N", "N", "E"
'                PersonaColLV 10, "N", "S", "", "N", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Set MeuLV.cmdconsulta(9).PictureNormal = MeuLV.ImageList1.ListImages(16).Picture
'            MeuLV.cmdconsulta(9).ToolTipText = "CD - Comunica��o de Desvio"
'            Set MeuLV.cmdconsulta(11).PictureNormal = MeuLV.ImageList1.ListImages(9).Picture
'            MeuLV.cmdconsulta(11).ToolTipText = "Abertura de Retrabalho"
'            Set MeuLV.cmdconsulta(12).PictureNormal = MeuLV.ImageList1.ListImages(14).Picture
'            MeuLV.cmdconsulta(12).ToolTipText = "Baixa Parcial de OS/Opera��o"
'            Exit Sub
'        'Controle de Desenhos
        ElseIf QualLV = 10 Then
'            Set chamaForm = New frmCD
'            Formulario = "CD - Controle de Desenhos"
'            LegendaExc = "CD - Controle de Desenhos" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                frmFiltro.frmPeriodo.Visible = True
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,CAST(c.fce AS VARCHAR(4)) + ' - ' + c.projeto AS FCE,b.desenho,b.revisao,a.quantidade,a.pesounit,(a.quantidade*a.pesounit) as pesototal,a.datarecebido,a.ptempo + ' ' + a.punidade,a.usuario,a.dataini,a.datafim,a.croqui,a.status,a.observacao,a.ativo,a.detalhista from " & _
'                                                      "tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto order by a.idcd desc"
'                If FiltroGeral = "Ativos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,CAST(c.fce AS VARCHAR(4)) + ' - ' + c.projeto AS FCE,b.desenho,b.revisao,a.quantidade,a.pesounit,(a.quantidade*a.pesounit) as pesototal,a.datarecebido,a.ptempo + ' ' + a.punidade,a.usuario,a.dataini,a.datafim,a.croqui,a.status,a.observacao,a.ativo,a.detalhista from " & _
'                                                      "tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto where a.ativo = 'S' order by a.idcd desc"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "select top " & LimiteLinhas & " a.idcd,CAST(c.fce AS VARCHAR(4)) + ' - ' + c.projeto AS FCE,b.desenho,b.revisao,a.quantidade,a.pesounit,(a.quantidade*a.pesounit) as pesototal,a.datarecebido,a.ptempo + ' ' + a.punidade,a.usuario,a.dataini,a.datafim,a.croqui,a.status,a.observacao,a.ativo,a.detalhista from " & _
'                                                      "tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto where a.ativo <> 'S' order by a.idcd desc"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Identificador", "FCE", "Desenho", "Rev.", "Quant.", "Peso Unit.", "Peso Total", "Recebido", "Previs�o Det.", "Usu�rio", "Data inicio", "Data fim", "Croqui", "Status", "Observa��o", "Ativo", "Detalhista", "", "", "", ""
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
'            Exit Sub
'        'F�rmula = Centro de Custo
        ElseIf QualLV = 11 Then
'            Set chamaForm = New frmFormulaCC
'            Formulario = "F�rmula - Centro de Custo"
'            LegendaExc = "F�rmula - Centro de Custo" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select a.CODREDUZIDO,a.NOME, 'formula' = case when max(b.nmform) IS NULL then '-' else 'com formula' end from CORPORERM.dbo.GCCUSTO as a left join Ferramentaria.dbo.tbFormula as b " & _
'                "on a.CODREDUZIDO = b.codreduzido COLLATE SQL_Latin1_General_CP1_CI_AS Where ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '1000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '3000' or ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '7000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '5000' or " & _
'                "ativo = 'T' and substring(a.CODREDUZIDO,1,4) = '7000' or ativo  = 'T' and substring(a.CODREDUZIDO,1,4) = '4000' group by a.ID,a.CODREDUZIDO,a.NOME order by a.CODREDUZIDO"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "Centro de Custo", "Nome Centro de Custo", "F�rmula", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "F�rmulas - Centro de Custo"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Exit Sub
        'QUALIDADE - RNCF (Registro de N�o Conformidade de Fabrica��o)
        ElseIf QualLV = 12 Then
'            Set chamaForm = New frmRNCF
'            Formulario = "RNCF - Registro de N�o Conformidade de Fabrica��o"
'            LegendaExc = "RNCF - Registro de N�o Conformidade de Fabrica��o" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'
'                If FiltroGeral = "Todos" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                      "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                If FiltroGeral = "CD - Comunica��o de Desvio" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                          "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                                                          "where a.status = 6 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                If FiltroGeral = "CODAC - Coleta de Dados e An�lise de Causas" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                          "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                                                          "where a.status = 7 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                If FiltroGeral = "DAAC - Defini��o da A��o e An�lise Concluida" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                           "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                                                           "where a.status = 8 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'
'                If FiltroGeral = "EVA - Execu��o e Verifica��o da A��o" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                           "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                                                           "where a.status = 9 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'                If FiltroGeral = "TAC - Tomada de A��o Concluida" Then SqlLV = "select a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20) as projeto,substring(a.observacao,1,100) as observacao,a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento from tbComunicacaoDesvio as a left join tbMPitens as b on a.idos = b.idos " & _
'                                                           "left join tbmp as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbRNC as e on a.idcd = e.idcd left join tbRetrabalho as h on a.idcd = h.idcd " & _
'                                                           "where a.status = 10 group by a.idcd,a.dataabertura,a.responsavel,a.idos,d.fce,substring(d.projeto,1,20),substring(a.observacao,1,100),a.status,e.idrnc,e.dataconclusao,e.gerouretrabalho,h.idretrabalho,e.datafechamento"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = True
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'            QtdColReal = 0
'            MontaCabLV "CD n�", "Data Abertura", "Respons�vel", "OS n�", "FCE", "Projeto", "Observa��o", "Status", "RNC n�", "Data Conclus�o", "Retrabalho", "Retrabalho n�", "Data Fechamento", "", "", "", "", "", "", "", ""
'            DimensionaLV "RNCF - Registro de N�o Conformidade de Fabrica��o"
'            MontaCabecalhoLV
'            MontaDadosLV "S" ' Zero a direita na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 3, "N", "N", "", "N", "S", "N", "E"
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
'
'            Exit Sub
        'USU�RIOS
        ElseIf QualLV = 13 Then
            Set chamaForm = New frmUsuarios
            Formulario = "Usu�rios"
            LegendaExc = "Usu�rio" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
 
                If FiltroGeral = "Todos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.codven,a.nomeven,a.codusuariototvs,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.codven,a.nomeven,a.codusuariototvs,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo = 'S'"
                If FiltroGeral = "N�o ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.codven,a.nomeven,a.codusuariototvs,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo is null or a.ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            
            QtdColReal = 0
            MontaCabLV "C�digo", "Nome do usu�rio", "Grupo", "TOTVS (codven)", "TOTVS (Nome)", "TOTVS (codusuario)", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Usu�rios"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Exit Sub
        'GRUPOS
        ElseIf QualLV = 14 Then
            Set chamaForm = New frmGrupos
            Formulario = "Grupos"
            LegendaExc = "Grupo" 'Usado na mensagem de exclus�o
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select * from tbgrupo"
                If FiltroGeral = "Ativos" Then SqlLV = "select * from tbgrupo where ativo = 'S'"
                If FiltroGeral = "N�o ativos" Then SqlLV = "select * from tbgrupo where ativo is null or ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(11).Visible = False
            MeuLV.cmdconsulta(12).Visible = False
            MeuLV.cmdconsulta(0).Visible = False
            QtdColReal = 0
            MontaCabLV "C�digo", "Nome do grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Grupos"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Exit Sub
        'OS FECHAMENTO - PERMISS�O DE COLABORADORES
        ElseIf QualLV = 15 Then
'            Set chamaForm = New frmPerColab
'            Formulario = "OS Fechamento - Permiss�o de Colaboradores"
'            LegendaExc = "Permiss�o do colaborador" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
'                                                      "a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 left join tbautfechaos as c on a.chapa = c.chapa COLLATE SQL_Latin1_General_CP1_CI_AS ORDER BY a.chapa"
'                If FiltroGeral = "Ativos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
'                                                       "a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 left join tbautfechaos as c on a.chapa = c.chapa COLLATE SQL_Latin1_General_CP1_CI_AS where  c.chapa is not null ORDER BY a.chapa"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "select a.CHAPA,b.NOME,case when c.chapa is not null then 'S' else 'N' end as ativo from CORPORERM.dbo.PFUNC as a inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and " & _
'                                                           "a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10 left join tbautfechaos as c on a.chapa = c.chapa COLLATE SQL_Latin1_General_CP1_CI_AS where  c.chapa is null or c.chapa = 'N' ORDER BY a.chapa"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(4).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'
'            QtdColReal = 0
'            MontaCabLV "Chapa", "Nome", "Permiss�o", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "OS Fechamento - Permiss�o de Colaboradores"
'            MontaCabecalhoLV
'            MontaDadosLV "N" 'Coloca zeros a esquerda na primeira coluna
'            If checaFiltro = True Then
'                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Exit Sub
        'LF - Libera��o de Fabrica��o
        ElseIf QualLV = 16 Then
'            Set chamaForm = New frmRelInsp
'            Formulario = "Relat�rios de Inspe��o"
'            LegendaExc = "Relat�rios de Inspe��o" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.codprojeto,a.fce,a.projeto,c.nome from tbProjetos as a inner join tbFO as b on a.fce=b.fce inner join tbclifor as c on b.codclifor = c.codclifor where a.fce > 2000 Order by a.fce desc,a.descricao"
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
'            MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Relat�rios de Inspe��o"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
''                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(19).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Relat�rio de Inspe��o - Fabrica��o"
'
'            Set MeuLV.cmdconsulta(6).PictureNormal = MeuLV.ImageList1.ListImages(21).Picture
'            MeuLV.cmdconsulta(6).ToolTipText = "Relat�rio de Inspe��o - Pintura"
'            Exit Sub
        'RO - Relat�rio de Expedi��o
        ElseIf QualLV = 17 Then
'            Set chamaForm = New frmRelExp
'            Formulario = "Relat�rios de Expedi��o"
'            LegendaExc = "Relat�rios de Expedi��o" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select top " & LimiteLinhas & " a.codprojeto,a.fce,a.projeto,c.nome from tbProjetos as a inner join tbFO as b on a.fce=b.fce inner join tbclifor as c on b.codclifor = c.codclifor where a.fce > 2000 Order by a.fce desc,a.descricao"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            MeuLV.ListView1.CheckBoxes = True
'            MeuLV.cmdconsulta(5).Visible = False
'            MeuLV.cmdconsulta(6).Visible = False
'            MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(11).Visible = False
'            MeuLV.cmdconsulta(12).Visible = False
'
'            QtdColReal = 0
'            MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
'            DimensionaLV "Relat�rios de Inspe��o"
'            MontaCabecalhoLV
'            MontaDadosLV "S"
'            If checaFiltro = True Then
''                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'
'            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(27).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Relat�rio de Expedi��o"
'
'            Exit Sub
        ElseIf QualLV = 18 Then
'            Set chamaForm = New frmADP
'            Formulario = "ADP"
'            LegendaExc = "ADP" 'Usado na mensagem de exclus�o
'            indiceVarGlobal = 1
'            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
'                MontaFiltro
'                If FiltroGeral = "" Then frmFiltro.Show 1
'                If MeuLV.Visible = True Then Unload MeuLV
'                If FiltroGeral = "Todos" Then SqlLV = "select b.codcolaborador,b.nomecolaborador,a.tipoadp,cast(a.dias as int),CONVERT (VARCHAR, a.datavencimento, 103) as datavencimento,CONVERT (VARCHAR, a.datadevolucao, 103) as datadevolucao,CONVERT (VARCHAR, a.dataavaliacao, 103) as dataavaliacao ,a.nota,a.statusimpressao,a.statusavaliacao,a.ativo,a.id from tbListaADP as a inner join tbColaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador=b.id and b.ativo = 'S' and b.tipo = 'colaborador' order by a.dias"
'                If FiltroGeral = "Ativos" Then SqlLV = "select b.codcolaborador,b.nomecolaborador,a.tipoadp,cast(a.dias as int),CONVERT (VARCHAR, a.datavencimento, 103) as datavencimento,CONVERT (VARCHAR, a.datadevolucao, 103) as datadevolucao,CONVERT (VARCHAR, a.dataavaliacao, 103) as dataavaliacao ,a.nota,a.statusimpressao,a.statusavaliacao,a.ativo,a.id from tbListaADP as a inner join tbColaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador=b.id and b.ativo = 'S' and b.tipo = 'colaborador' where a.ativo is not null order by a.dias"
'                If FiltroGeral = "N�o ativos" Then SqlLV = "select b.codcolaborador,b.nomecolaborador,a.tipoadp,cast(a.dias as int),CONVERT (VARCHAR, a.datavencimento, 103) as datavencimento,CONVERT (VARCHAR, a.datadevolucao, 103) as datadevolucao,CONVERT (VARCHAR, a.dataavaliacao, 103) as dataavaliacao ,a.nota,a.statusimpressao,a.statusavaliacao,a.ativo,a.id from tbListaADP as a inner join tbColaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador=b.id and b.ativo = 'S' and b.tipo = 'colaborador' where a.ativo is null order by a.dias"
'            Else
'                If MeuLV.Visible = True Then Unload MeuLV
'            End If
'            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
'            MeuLV.cmdconsulta(10).Visible = True
'
'            QtdColReal = 0
'            MontaCabLV "Registro", "Nome", "Tipo", "Periodo", "Vencimento", "Devolu��o", "Avaliado em", "Pontua��o", "Impresso", "Status ADP", "Ativo", "id", "", "", "", ""
'            DimensionaLV "ADP - Avalia��o de Desempenho Profissional"
'            MontaCabecalhoLV
'            MontaDadosLV "N"
'            If checaFiltro = True Then
'                PersonaColLV 3, "N", "N", "", "N", "N", "N", "D"
'                PersonaColLV 7, "S", "S", "%", "N", "N", "S", "D"
'                PersonaColLV 8, "N", "N", "", "S", "N", "N", "E"
'                PersonaColLV 10, "N", "N", "", "S", "N", "N", "E"
'            End If
'            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
'            MeuLV.Label2.Caption = FiltroGeral
'            CompoeComboLV MeuLV.Combo1
'            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(5).Picture
'            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(6).Picture
'            MeuLV.cmdconsulta(4).ToolTipText = "Avaliar"
'            MeuLV.cmdconsulta(5).ToolTipText = "Sair"
'            MeuLV.cmdconsulta(6).Visible = False
'            MeuLV.cmdconsulta(7).Visible = False
'            Exit Sub
        End If
        Set frmFiltro = Nothing
        Set MeuLV = Nothing
        Set chamaForm = Nothing
TrataErro:
    If Err.Number = 400 Then
        FiltroGeral = "Ativos"
        Resume Next
    End If
End Sub

Public Sub carregaTABS(vTab1 As String, vTab2 As String, vTab3 As String, vTab4 As String, vTab5 As String, vTab6 As String, vTab7 As String, vTab8 As String, vTab9 As String, vTab10 As String, vTab11 As String, vTab12 As String, vTab13 As String, vTab14 As String, vTab15 As String)
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
    vTabela11 = vTab11
    vTabela12 = vTab12
    vTabela13 = vTab13
    vTabela14 = vTab14
    vTabela15 = vTab15
End Sub

Public Sub CarregaSQLExcluir(QLV As Integer)
    Dim rsExcLVGeral As New ADODB.Recordset
    Dim P As Integer
    If QLV = 0 Then
        'frmDemitirColaborador.Show 1
        'gravaLog varGlobal, MeuLV.ListView1.SelectedItem.ListSubItems.Item(1), "-"
    ElseIf QLV = 1 Then
        'SqlExcLVGeral = "Delete from tbColaboradores where a.codcoligada = '" & vCodcoligada & "' and cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresesc where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradorescur where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresexp where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoreshist where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'"
    ElseIf QLV = 2 Then
        'SqlExcLVGeral = "Delete from tbDepartamentos where codDepartamento= '" & Val(varGlobal) & "' ;Delete from tbDepartamentosHistResp where codDepartamento= '" & Val(varGlobal) & "'"
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                Msgbox varGlobal
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbDepartamentos set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and coddepartamento = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        Else
            cnBanco.CommitTrans
            Exit Sub
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 3 Then
        'SqlExcLVGeral = "Delete from tbSetores where codSetor= '" & Val(varGlobal) & "' ;Delete from tbSetoresHistResp where codSetor= '" & Val(varGlobal) & "'"
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbSetores set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codsetor = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 4 Then
        'NAO EXCLUI O PRODUTO, EXCLUI OS DADOS DAS F�RMULAS REFERENTE AO PRODUTO
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o da " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "Delete from tbmateriais where idprd = '" & Val(MeuLV.ListView1.ListItems.Item(P)) & "'"
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                    
                    SqlExcLVGeral = "Delete from tbConstantes where idprd = '" & Val(MeuLV.ListView1.ListItems.Item(P)) & "'"
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                    
                End If
            Next
        End If
        cnBanco.CommitTrans
        
    ElseIf QLV = 5 Then
        'SqlExcLVGeral = "Delete from tbHabilidades where codHabilidade= " & Val(varGlobal)
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbHabilidades set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codhabilidade = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 6 Then
        'SqlExcLVGeral = "Delete from tbEscolaridade where codescolaridade= " & Val(varGlobal)
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbEscolaridade set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codescolaridade = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 7 Then
        SqlExcLVGeral = "Delete from tbdesenhos where codcoligada = '" & vCodcoligada & "' and iddesenho= '" & Val(varGlobal) & "' ;Delete from tbdesenhos where codcoligada = '" & vCodcoligada & "' and iddesenho= '" & Val(varGlobal) & "'"
        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
    ElseIf QLV = 8 Then
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            SqlExcLVGeral = "Select count(*) from tbItemLM as a where a.fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and a.codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
            If rsExcLVGeral.Fields(0) > 0 Then ' Desativa
                rsExcLVGeral.Close
                SqlExcLVGeral = "delete from tbItemLM where fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                mobjMsg.Abrir "Curso/treinamento DESATIVADO com sucesso", Ok, informacao, "IMRM"
            End If
            rsExcLVGeral.Close
        
            SqlExcLVGeral = "Select count(*) from tbLM as a where a.fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and a.codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
            If rsExcLVGeral.Fields(0) > 0 Then ' Desativa
                rsExcLVGeral.Close
                SqlExcLVGeral = "delete from tbLM where fce = '" & Val(Mid$(varGlobal, 1, 4)) & "' and codlm = '" & Val(Mid$(varGlobal, 5, 6)) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                mobjMsg.Abrir "LM Excluida com sucesso", Ok, informacao, "IMRM"
            End If
            'rsExcLVGeral.Close
            Set rsExcLVGeral = Nothing
        End If
        cnBanco.CommitTrans
        'rsExcLVGeral.Close
        'SqlExcLVGeral = "Delete from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codTreinamento=  '" & Val(varGlobal) & "' ;Delete from tbTreinamentosRev where codcoligada = '" & vCodcoligada & "' and codTreinamento= '" & Val(varGlobal) & "'"
    ElseIf QLV = 9 Then
        Dim vPlanej As Integer, vOS As Integer
        vPlanej = Val(Mid$(varGlobal, 1, 6))
        vOS = Val(Mid$(varGlobal, 7, 6))
        If vOS = 0 Then
            SqlExcLVGeral = "Delete from tbmp where idprogramacao = '" & vPlanej & "' ;Delete from tbMPItens where idprogramacao = '" & vPlanej & "' ;Delete from tbositens where idprogramacao = '" & vPlanej & "' ;Delete from tbos where idos = 0"
            rsExcLVGeral.Open SqlExcLVGeral, cnBanco
            mobjMsg.Abrir "Registro excluido com sucesso", Ok, informacao, "IMRM"
        Else
            mobjMsg.Abrir "Registro n�o pode ser excluido", Ok, critico, "IMRM"
        End If
        
        'SqlExcLVGeral = "Select a.codmatriz from tbmatriz as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = b.codmatriz where a.codmatriz = '" & Val(varGlobal) & "'"
        'rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
        'If rsExcLVGeral.RecordCount = 0 Then
        '    rsExcLVGeral.Close
        '    Set rsExcLVGeral = Nothing
        '    mobjMsg.Abrir "Confirma exclus�o da " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        '    If Tp = 1 Then
        '        SqlExcLVGeral = "Delete from tbMatriz where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizCur where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizEsc where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizExp where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizHab where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "'"
        '        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
        '        mobjMsg.Abrir "Registro excluido com sucesso", Ok, informacao, "IMRM"
        '    End If
        'Else
        '    rsExcLVGeral.Close
        '    Set rsExcLVGeral = Nothing
        '    mobjMsg.Abrir "Matriz n�o poder ser excluida! A Chave prim�ria est� sendo utilizada em outras tabelas", Ok, critico, "Aten��o"
        'End If
    ElseIf QLV = 10 Then
        Dim strResultado As String
        mobjMsg.Abrir "Confirma o Cancelamento da " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    'If strResultado <> "" Then
                        SqlExcLVGeral = "UPDATE tbCD set ativo = 'N' where idcd = '" & Val(varGlobal) & "'"
                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                    'Else
                    '    MsgBox "� necess�rio justificar o cancelamento"
                    'End If
                End If
            Next
            mobjMsg.Abrir "Cancelamento realizado!", Ok, critico, "Aten��o"
        End If
    ElseIf QLV = 11 Then
        'SqlExcLVGeral = "Delete from tbAvaliacao where codavaliacao= " & Val(varGlobal)
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbAvaliacao set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codavaliacao = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 15 Then
        'Ferramentaria - Exclui Autorizados a Fechar OS - Ordem de Servi�o
        'Dim strResultado As String
        mobjMsg.Abrir "Confirma o Cancelamento da " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    'If strResultado <> "" Then
                        SqlExcLVGeral = "delete from tbAutCCusto where chapa = '" & MeuLV.ListView1.ListItems.Item(P) & "'"
                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                        
                        SqlExcLVGeral = "delete from tbAutFechaOs where chapa = '" & MeuLV.ListView1.ListItems.Item(P) & "'"
                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
            mobjMsg.Abrir "Cancelamento realizado!", Ok, critico, "Aten��o"
        End If
    ElseIf QLV = 16 Then
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclus�o do " & LegendaExc & " selecionado?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            frmExcluiINTD.Show 1
        End If
        cnBanco.CommitTrans
    End If
End Sub

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

'Gera Avalia��o de Desempenho Profissional por colaborador
Public Function carregaADP()
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
End Function

Public Function montaDadosADP()
    Dim rsMontaDadosADP As New ADODB.Recordset
    Dim SqlMontaDadosADP As String
    
    Dim rsVerificaADP As New ADODB.Recordset
    Dim SqlVerificaADP As String
    Dim diasProximaADP As Integer
    
    'Todos os colaboradors com a quantidade de dias que est�o na matriz
    SqlMontaDadosADP = "select a.id, a.nomecolaborador, b.codmatriz, b.data, DATEDIFF(DAY,b.data,GETDATE()) from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf where a.ativo = 'S' and b.ativo = 'S'"
    rsMontaDadosADP.Open SqlMontaDadosADP, cnBanco, adOpenKeyset, adLockReadOnly
    For X = 1 To rsMontaDadosADP.RecordCount
        SqlVerificaADP = "Select * from tblistaADP where codcoligada = '" & vCodcoligada & "' and codcolaborador = '" & rsMontaDadosADP.Fields(0) & "' and statusavaliacao is null or codcoligada = '" & vCodcoligada & "' and codcolaborador = '" & rsMontaDadosADP.Fields(0) & "' and statusavaliacao <> 'Concluido'"
        rsVerificaADP.Open SqlVerificaADP, cnBanco, adOpenKeyset, adLockOptimistic
        'SE FOR = 0 NAO EXISTE AVALIACAO EM ABERTO PARA O COLABORADOR
        'ENTRA NA CONDI��O ABAIXO
        If rsVerificaADP.RecordCount = 0 Then
            diasTrabalhados = rsMontaDadosADP.Fields(4)
            avaliarAKDA = achaDias(rsMontaDadosADP.Fields(0))
            If Val(diasTrabalhados) > Val(avaliarAKDA) Then
                diasProximaADP = Val(diasTrabalhados / avaliarAKDA) * avaliarAKDA + avaliarAKDA
            Else
                diasProximaADP = avaliarAKDA
            End If
            ' AKI CHAMA ROTINA DE GRAVA��O
            rsVerificaADP.AddNew
            rsVerificaADP.Fields(1) = rsMontaDadosADP.Fields(0)
            rsVerificaADP.Fields(2) = tipoADP
            rsVerificaADP.Fields(3) = avaliarAKDA
            'Teste para corrigir o erro de 1 dia na avalia��o de desempenho
            rsVerificaADP.Fields(5) = rsMontaDadosADP.Fields(3) + (diasProximaADP - 1) 'Teste para corrigir o erro de 1 dia na avalia��o de desempenho
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
End Function

Public Function achaDias(vCodColab As String)
    Dim rsAchaDias As New ADODB.Recordset
    Dim SqlAchaDias As String
    Dim X As Integer
    
    achaDias = 0
    
    SqlAchaDias = "select a.id,a.codcolaborador,a.dias,a.statusavaliacao from tbListaADP as a where a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador = '" & vCodColab & "' and statusavaliacao = 'concluido' order by a.dias desc"
    rsAchaDias.Open SqlAchaDias, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsAchaDias.EOF Then achaDias = rsAchaDias.Fields(2)
    rsAchaDias.Close
    Set rsAchaDias = Nothing
    '--> SE ENCONTRAR AVALIA��ES JA CONCLUIDAS NO SISTEMA
    If achaDias > 0 Then
        For X = 0 To 10
            If vADP(X, 0) = "" Then Exit Function
            If Val(vADP(X, 0)) > achaDias Then
                achaDias = vADP(X, 0)
                tipoADP = vADP(X, 1)
                If diasTrabalhados < achaDias Then Exit Function
            End If
        Next
    '--> SE N�O ENCONTRAR AVALIA��ES CONCLUIDAS NO SISTEMA
    Else
        For X = 0 To 10
            If vADP(X, 0) = "" Then Exit Function
            achaDias = vADP(X, 0)
            tipoADP = vADP(X, 1)
            If diasTrabalhados < achaDias Then Exit Function
        Next
    End If
End Function
