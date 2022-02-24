Attribute VB_Name = "CompoeLVs"
Public apontaLV As Integer
Public indiceVarGlobal As Integer 'quantas colunas vai ter a variavel global
Public checaFiltro As Boolean
Public vADP(10, 1) As String
Public diasTrabalhados As Integer
Public avaliarAKDA As Integer
Public tipoADP As String

Public Sub MontaLV(QualLV As Integer)
        'COLABORADORES
        If QualLV = 0 Then
            If vIntegra = "S" Then buscaDemitidos
            If vCalcExp = "S" Then caculaTmpExp
            Set chamaForm = New frmColaboradores
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Colaboradores"
            LegendaExc = "Colaborador" 'Usado na mensagem de exclusão
            indiceVarGlobal = 2
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then
                    Unload MeuLV
                End If
                If FiltroGeral = "Todos" Then SqlLV = "Select a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'colaborador' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.ativo='S'"
                If FiltroGeral = "Ativos" Then SqlLV = "Select a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'colaborador' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where a.ativo = 'S' and b.ativo='S'"
                'If FiltroGeral = "Não ativos" Then SqlLV = "Select a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a left join tbcolaboradoresHist as b on a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'colaborador' left join tbMatriz as c on c.codmatriz=b.codmatriz left join tbcargos as d on d.codcargo = c.codcargo where a.ativo is null and b.ativo='S' and a.homologacaonum is null or a.ativo='N' and b.ativo='S' and a.homologacaonum is null"
                If FiltroGeral = "Não ativos" Then SqlLV = "Select a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,MAX(d.nomecargo) as nomecargo from tbcolaboradores as a left join tbcolaboradoresHist as b on a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'colaborador' left join tbMatriz as c on c.codmatriz=b.codmatriz left join tbcargos as d on d.codcargo = c.codcargo where a.codcoligada = '" & vCodcoligada & "' and a.ativo is null and a.homologacaonum is null or a.codcoligada = '" & vCodcoligada & "' and a.ativo='N' and a.homologacaonum is null group by a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav"
                If FiltroGeral = "Demitidos" Then SqlLV = "Select a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'colaborador' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where a.ativo is null  and b.ativo='S' and a.homologacaonum is not null or a.ativo='N' and b.ativo='S' and a.homologacaonum is not null"
                If FiltroGeral = "Afastados" Then SqlLV = "Select a.cpf,a.codcolaborador,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a left join tbcolaboradoresHist as b on a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'colaborador' left join tbMatriz as c on c.codmatriz=b.codmatriz left join tbcargos as d on d.codcargo = c.codcargo where a.codcoligada = '" & vCodcoligada & "' and a.ativo='A' and b.ativo = 'S' and a.homologacaonum is null"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            MeuLV.cmdconsulta(11).Visible = True
            MeuLV.cmdconsulta(12).Visible = True
            If apontaLV = 0 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            If apontaLV = 0 Then MeuLV.cmdconsulta(10).Visible = True Else MeuLV.cmdconsulta(10).Visible = False
            QtdColReal = 0
            MontaCabLV "Colaborador CPF", "Registro", "Nome", "CTPS nº", "Série", "Pontuação", "Ativo", "C. Avaliadas", "Cargos", "", "", "", "", "", "", ""
            DimensionaLV "Colaboradores"
            MontaCabecalhoLV
            MontaDadosLV "N"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "N", "N", "D"
                PersonaColLV 5, "S", "S", "%", "N", "N", "S", "D"
                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Set MeuLV.cmdconsulta(6).PictureNormal = MeuLV.ImageList1.ListImages(4).Picture
            MeuLV.cmdconsulta(6).ToolTipText = "Demitir colaborador"
            MeuLV.cmdconsulta(9).ToolTipText = "Admitir colaborador"
            If vIntegra = "S" Then
                MeuLV.cmdconsulta(6).UseGreyscale = True
                MeuLV.cmdconsulta(6).DragMode = 1
                MeuLV.cmdconsulta(6).SpecialEffect = cbEngraved
            End If

            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'CANDIDATOS
        ElseIf QualLV = 1 Then
            Set chamaForm = New frmCandidatos
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Candidatos"
            LegendaExc = "Candidato" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbColaboradores where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresesc where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradorescur where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresexp where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoreshist where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'"
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "Select a.cpf,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'candidato' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.ativo='S'"
                If FiltroGeral = "Ativos" Then SqlLV = "Select a.cpf,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'candidato' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where a.ativo='S' and b.ativo='S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "Select a.cpf,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'candidato' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where a.ativo is null  and b.ativo='S' or a.ativo='N' and b.ativo='S'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(10).Visible = True Else MeuLV.cmdconsulta(10).Visible = False
            
            QtdColReal = 0
            MontaCabLV "Candidatos CPF", "Nome", "CTPS nº", "Série", "Pontuação", "Ativo", "C. Avaliadas", "Cargos", "", "", "", "", "", "", "", ""
            DimensionaLV "Candidatos"
            MontaCabecalhoLV
            MontaDadosLV "N"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "N", "N", "D"
                PersonaColLV 4, "S", "S", "%", "N", "N", "S", "D"
                PersonaColLV 5, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'DEPARTAMENTOS
        ElseIf QualLV = 2 Then
            'Unload MeuLV
            Set chamaForm = New frmDepartamentos
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Departamentos"
            LegendaExc = "Departamento" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbDepartamentos where codDepartamento= '" & Val(varGlobal) & "' ;Delete from tbDepartamentosHistResp where codDepartamento= '" & Val(varGlobal) & "'"
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select a.coddepartamento,a.nomedepartamento,c.nomecolaborador,a.ativo from tbdepartamentos as a left join tbDepartamentosHistResp as b on a.coddepartamento = b.coddepartamento left join tbcolaboradores as c on  b.codcolaborador = c.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and b.datafim is null"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.coddepartamento,a.nomedepartamento,c.nomecolaborador,a.ativo from tbdepartamentos as a left join tbDepartamentosHistResp as b on a.coddepartamento = b.coddepartamento left join tbcolaboradores as c on  b.codcolaborador = c.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and b.datafim is null and a.ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select a.coddepartamento,a.nomedepartamento,c.nomecolaborador,a.ativo from tbdepartamentos as a left join tbDepartamentosHistResp as b on a.coddepartamento = b.coddepartamento left join tbcolaboradores as c on  b.codcolaborador = c.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and b.datafim is null and a.ativo is null or a.codcoligada = '" & vCodcoligada & "' and a.ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            
            If apontaLV = 2 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Departamento", "Responsável", "Ativo", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Departamentos"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'SETORES
        ElseIf QualLV = 3 Then
            Set chamaForm = New frmSetores
'            frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Setores"
            LegendaExc = "Setor" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbSetores where codSetor= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select a.codsetor,a.nomesetor,b.nomedepartamento,d.nomecolaborador,a.ativo from tbSetores as a left join tbdepartamentos as b on a.codcoligada = b.codcoligada and a.coddepartamento = b.coddepartamento left join tbsetoresHistResp as c on a.codcoligada = c.codcoligada and a.coddepartamento = c.coddepartamento and a.codsetor = c.codsetor left join tbcolaboradores as d on a.codcoligada = d.codcoligada and c.codcolaborador = d.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and c.datafim is null"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.codsetor,a.nomesetor,b.nomedepartamento,d.nomecolaborador,a.ativo from tbSetores as a left join tbdepartamentos as b on a.codcoligada = b.codcoligada and a.coddepartamento = b.coddepartamento left join tbsetoresHistResp as c on a.codcoligada = c.codcoligada and a.coddepartamento = c.coddepartamento and a.codsetor = c.codsetor left join tbcolaboradores as d on a.codcoligada = d.codcoligada and c.codcolaborador = d.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and a.ativo = 'S' and c.datafim is null"
                If FiltroGeral = "Não ativos" Then SqlLV = "select a.codsetor,a.nomesetor,b.nomedepartamento,d.nomecolaborador,a.ativo from tbSetores as a left join tbdepartamentos as b on a.codcoligada = b.codcoligada and a.coddepartamento = b.coddepartamento left join tbsetoresHistResp as c on a.codcoligada = c.codcoligada and a.coddepartamento = c.coddepartamento and a.codsetor = c.codsetor left join tbcolaboradores as d on a.codcoligada = d.codcoligada and c.codcolaborador = d.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and a.ativo is null and c.datafim is null or a.codcoligada = '" & vCodcoligada & "' and a.ativo ='N' and c.datafim is null"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 3 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Setor", "Departamento", "Responsável", "Ativo", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Setores"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'CARGOS
        ElseIf QualLV = 4 Then
            Set chamaForm = New frmCargos
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Cargos"
            LegendaExc = "Cargo" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbcargos where codcargo= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select codcargo,nomecargo,codcbo,ativo,descricao from tbcargos where codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select codcargo,nomecargo,codcbo,ativo,descricao from tbcargos where codcoligada = '" & vCodcoligada & "' and ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select codcargo,nomecargo,codcbo,ativo,descricao from tbcargos where codcoligada = '" & vCodcoligada & "' and ativo is null or codcoligada = '" & vCodcoligada & "' and ativo = 'N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 4 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 4 Then MeuLV.cmdconsulta(10).Visible = True Else MeuLV.cmdconsulta(10).Visible = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Cargo", "CBO nº", "Ativo", "Descrição", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Cargos"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'HABILIDADES
        ElseIf QualLV = 5 Then
            Set chamaForm = New frmHabilidades
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Habilidades"
            LegendaExc = "Habilidade" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbHabilidades where codHabilidade= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select codhabilidade,nomehabilidade,peso,ativo from tbHabilidades where codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select codhabilidade,nomehabilidade,peso,ativo from tbHabilidades where codcoligada = '" & vCodcoligada & "' and ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select codhabilidade,nomehabilidade,peso,ativo from tbHabilidades where codcoligada = '" & vCodcoligada & "' and ativo is null or codcoligada = '" & vCodcoligada & "' and ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 5 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Habilidade", "Peso", "Ativo", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Habilidades"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 2, "N", "N", "", "N", "N", "N", "D"
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'ESCOLARES
        ElseIf QualLV = 6 Then
            Set chamaForm = New frmEscolaridade
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Escolaridade"
            LegendaExc = "Escolaridade" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbEscolaridade where codescolaridade= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select codescolaridade,nomeescolaridade,peso,ativo from tbEscolaridade where codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select codescolaridade,nomeescolaridade,peso,ativo from tbEscolaridade where codcoligada = '" & vCodcoligada & "' and ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select codescolaridade,nomeescolaridade,peso,ativo from tbEscolaridade where codcoligada = '" & vCodcoligada & "' and ativo is null or codcoligada = '" & vCodcoligada & "' ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 6 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Escolaridade", "Peso", "Ativo", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Escolaridades"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 2, "N", "N", "", "N", "N", "N", "D"
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'REQUISICOES
        ElseIf QualLV = 7 Then
            Set chamaForm = New frmRequisicao
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Requisição"
            LegendaExc = "Requisição" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbRequisicoes where codrequisicao= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select codrequisicao,datarequisicao,origem,nomerequisitante,ativo from  tbrequisicoes where codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select codrequisicao,datarequisicao,origem,nomerequisitante,ativo from  tbrequisicoes where codcoligada = '" & vCodcoligada & "' and ativo='S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select codrequisicao,datarequisicao,origem,nomerequisitante,ativo from  tbrequisicoes where codcoligada = '" & vCodcoligada & "' and ativo='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Código", "Data requisição", "Origem", "Requisitante", "Ativo", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Requisições"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'TREINAMENTOS
        ElseIf QualLV = 8 Then
            Set chamaForm = New frmTreinamentos
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Treinamentos"
            LegendaExc = "Treinamento" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbTreinamentos where codTreinamento= "
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select codtreinamento,nometreinamento,origem,introdutorio,obrigatorio,tipo,ativo from tbTreinamentos where codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select codtreinamento,nometreinamento,origem,introdutorio,obrigatorio,tipo,ativo from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select codtreinamento,nometreinamento,origem,introdutorio,obrigatorio,tipo,ativo from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo is null or codcoligada = '" & vCodcoligada & "' and ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Código", "Treinamento", "Origem", "Introdutório", "Obrigatório", "Tipo", "Ativo", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Treinamentos"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'MATRIZES
        ElseIf QualLV = 9 Then
            Set chamaForm = New frmMatrizCapacitacao
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Matriz"
            LegendaExc = "Matriz" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbMatriz where codMatriz= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select tbMatriz.codmatriz,tbMatriz.codcargo,tbcargos.nomecargo,tbMatriz.nivel,tbMatriz.atividades,tbMatriz.ativo from tbMatriz inner join tbcargos on tbmatriz.codcoligada = tbcargos.codcoligada where tbMatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.codcargo = tbCargos.codcargo order by tbMatriz.codmatriz"
                If FiltroGeral = "Ativos" Then SqlLV = "select tbMatriz.codmatriz,tbMatriz.codcargo,tbcargos.nomecargo,tbMatriz.nivel,tbMatriz.atividades,tbMatriz.ativo from tbMatriz inner join tbcargos on tbmatriz.codcoligada = tbcargos.codcoligada where tbMatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.codcargo = tbCargos.codcargo and tbMatriz.ativo = 'S' order by tbMatriz.codmatriz"
                If FiltroGeral = "Não ativos" Then SqlLV = "select tbMatriz.codmatriz,tbMatriz.codcargo,tbcargos.nomecargo,tbMatriz.nivel,tbMatriz.atividades,tbMatriz.ativo from tbMatriz inner join tbcargos on tbmatriz.codcoligada = tbcargos.codcoligada where tbMatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.codcargo = tbCargos.codcargo and tbMatriz.ativo = 'N' or tbMatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.ativo is null order by tbMatriz.codmatriz"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            If apontaLV = 9 Then MeuLV.cmdconsulta(10).Visible = True Else MeuLV.cmdconsulta(10).Visible = False
            QtdColReal = 0
            MontaCabLV "Matriz", "Cód. Cargo", "Nome Cargo", "Nível", "Atividades", "Ativo", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Matrizes"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 1, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 5, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'PROGRAMAÇÃO
        ElseIf QualLV = 10 Then
            Set chamaForm = New frmProgramacao
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Programação"
            LegendaExc = "Programação" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbMatriz where codMatriz= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                frmFiltro.frmPeriodo.Visible = True
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                      "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where b.ativo = 'S' and h.codusuario = '" & CodUsu & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                       "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where b.ativo = 'S' and a.ativo = 'S' and h.codusuario ='" & CodUsu & "'"
                If FiltroGeral = "Ativos pendentes" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                                 "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where b.ativo = 'S' and a.ativo = 'S' and a.status='Pendente' and h.codusuario ='" & CodUsu & "'"
                'FILTRA COM DATA
                If FiltroGeral = "Ativos agendados" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                                 "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where " & _
                                                                 "b.ativo = 'S' and a.ativo = 'S' and a.status='Agendado' and g.avaldata BETWEEN '" & dataFilter1 & "' AND  '" & dataFilter2 & "' and h.codusuario ='" & CodUsu & "' or b.ativo = 'S' and a.ativo = 'S' and a.status='Agendado' and g.avaldata is null and h.codusuario ='" & CodUsu & "' or " & _
                                                                 "b.ativo = 'S' and a.ativo = 'S' and a.status='Reagendado' and g.avaldata BETWEEN '" & dataFilter1 & "' AND '" & dataFilter2 & "' and h.codusuario ='" & CodUsu & "' or b.ativo = 'S' and a.ativo = 'S' and a.status='Reagendado' and g.avaldata is null and h.codusuario ='" & CodUsu & "'"
                'FILTRA COM DATA
                If FiltroGeral = "Ativos concluidos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                                  "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where " & _
                                                                  "b.ativo = 'S' and a.ativo = 'S' and a.status='Concluido' and g.avaldata BETWEEN '" & dataFilter1 & "' AND '" & dataFilter2 & "' and h.codusuario ='" & CodUsu & "' or " & _
                                                                  "b.ativo = 'S' and a.ativo = 'S' and a.status='Concluido' and g.avaldata is null and h.codusuario ='" & CodUsu & "'"
                If FiltroGeral = "Ativos desmarcados" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                                   "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where b.ativo = 'S' and a.ativo = 'S' and a.status='Desmarcado' and h.codusuario ='" & CodUsu & "'"
                If FiltroGeral = "Cancelados" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao,g.avaldata,e.idGrFase,i.nomesetor from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                           "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where b.ativo = 'S' and a.ativo = 'N' and a.status='Cancelado' and h.codusuario ='" & CodUsu & "'"
                'If FiltroGeral = "Não ativos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,a.ativo,a.status,a.tipoprogramacao from tbPendentesCur as a inner join tbcolaboradores as b on b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel where a.ativo = 'N' or a.ativo is null"
                If FiltroGeral = "Programação" Then SqlLV = "select SUBSTRING(max(a.cpf),0,1),SUBSTRING(max(b.nomecolaborador),0,1),str(max(a.codmatriz),0,1) as codmatriz,SUBSTRING(max(d.nomecargo),0,1),max(a.codtreinamento) as codtreinamento,e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')',a.codprogramacao,max(a.ativo) as ativo,SUBSTRING(max(a.status),0,1) as status,str(max(a.tipoprogramacao),0,1),max(g.avaldata),e.idGrFase,max(i.nomesetor) from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel left join tbprogramacao as g on a.codprogramacao = g.codprogramacao " & _
                                                            "inner join tbUsuMultiplic as h on a.codtreinamento = h.codtreinamento inner join tbSetores as i on c.codsetor = i.codsetor where a.codprogramacao is not null and h.codusuario ='" & CodUsu & "' group by a.codprogramacao, e.nometreinamento + ' (' + isnull(f.nomenivel,'-') + ')'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 10 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            If apontaLV = 10 Then MeuLV.cmdconsulta(10).Visible = True Else MeuLV.cmdconsulta(10).Visible = False
            QtdColReal = 0
            MontaCabLV "Colaboradores CPF", "Nome Colaborador", "Matriz", "Cargo", "Cod. treinamento", "Nome treinamento", "Programação", "Ativo", "Status", "Tipo", "Av. Eficácia", "Fase", "Setor", "", "", ""
            DimensionaLV "Programações"
            MontaCabecalhoLV
            MontaDadosLV "N"
            If checaFiltro = True Then
                PersonaColLV 2, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 4, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 6, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 7, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.cmdconsulta(6).ToolTipText = "Cancelar treinamento"
            MeuLV.Label2.Caption = FiltroGeral
            MeuLV.Label4.Caption = Format(dataFilter1, "dd/mm/yyyy") & " - " & Format(dataFilter2, "dd/mm/yyyy")
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'AVALIAÇÃO
        ElseIf QualLV = 11 Then
            Set chamaForm = New frmAvaliacoes
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Avaliação"
            LegendaExc = "Avaliação" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbMatriz where codMatriz= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select a.codavaliacao,a.nomeavaliacao,a.tipo,a.peso,a.ativo from tbAvaliacao as a where a.codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.codavaliacao,a.nomeavaliacao,a.tipo,a.peso,a.ativo from tbAvaliacao as a where a.codcoligada = '" & vCodcoligada & "' and a.ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select a.codavaliacao,a.nomeavaliacao,a.tipo,a.peso,a.ativo from tbAvaliacao as a where a.codcoligada = '" & vCodcoligada & "' and a.ativo is null or a.codcoligada = '" & vCodcoligada & "' and ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 11 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            'If apontaLV = 10 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Nome avaliação", "Tipo", "Peso", "Ativo", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Avaliações"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 3, "N", "N", "", "N", "N", "N", "D"
                PersonaColLV 4, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'RESTRIÇÕES
        ElseIf QualLV = 12 Then
            Set chamaForm = New frmRecapacitacao
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Reprovados"
            LegendaExc = "Reprovados" 'Usado na mensagem de exclusão
            'SqlExcLVGeral = "Delete from tbMatriz where codMatriz= " & Val(varGlobal)
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento+ ' ('+nomenivel+')',a.codprogramacao,a.nota,a.situacao,a.ativo,a.observacao from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada ='" & vCodcoligada & "' and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and a.codnivel = f.codnivel where a.status='Concluido' and a.situacao='Reprovado' or a.status='Concluido' and a.situacao='Aprovado com restrição' or a.status= 'Recapacitação' and a.situacao='Reprovado'"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento+ ' ('+nomenivel+')',a.codprogramacao,a.nota,a.situacao,a.ativo,a.observacao,a.status " & _
                                                       "from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = 1 and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d " & _
                                                       "on d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and " & _
                                                       "a.codnivel = f.codnivel where a.ativo = 'S' and a.status='Concluido' and a.situacao='Reprovado' or a.status='Concluido' and a.situacao='Aprovado com restrição'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select a.cpf,b.nomecolaborador,a.codmatriz,d.nomecargo,a.codtreinamento,e.nometreinamento+ ' ('+nomenivel+')',a.codprogramacao,a.nota,a.situacao,a.ativo,a.observacao,a.status " & _
                                                           "from tbPendentesCur as a inner join tbcolaboradores as b on a.codcoligada = 1 and b.cpf = a.cpf inner join tbmatriz as c on c.codmatriz = a.codmatriz inner join tbcargos as d on " & _
                                                           "d.codcargo = c.codcargo inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento left join tbTreinamentosNiv as f on a.codtreinamento = f.codtreinamento and " & _
                                                           "a.codnivel = f.codnivel where a.ativo = 'N' and a.status='Concluido' and a.situacao='Reprovado' or a.ativo is null and a.status='Concluido' and a.situacao='Reprovado' or " & _
                                                           "a.ativo = 'N' and a.status='Concluido' and a.situacao='Aprovado com restrição' or a.ativo is null and a.status='Concluido' and a.situacao='Aprovado com restrição' or a.status = 'Recapacitação'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 12 Then MeuLV.ListView1.CheckBoxes = True Else MeuLV.ListView1.CheckBoxes = False
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Colaboradores CPF", "Nome Colaborador", "Matriz", "Cargo", "Cod. treinamento", "Nome treinamento", "Programação", "Pontuação", "Situação", "Ativo", "Observação", "", "", "", "", ""
            DimensionaLV "Reprovados"
            MontaCabecalhoLV
            MontaDadosLV "N"
            If checaFiltro = True Then
                PersonaColLV 2, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 4, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 6, "N", "N", "", "N", "S", "N", "E"
                PersonaColLV 7, "S", "S", "%", "N", "N", "S", "E"
                PersonaColLV 9, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(1).Picture
            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(2).Picture
            MeuLV.cmdconsulta(4).ToolTipText = "Aprovar"
            MeuLV.cmdconsulta(5).ToolTipText = "Recapacitação"
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'USUÁRIOS
        ElseIf QualLV = 13 Then
            Set chamaForm = New frmUsuarios
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Usuários"
            LegendaExc = "Usuário" 'Usado na mensagem de exclusão
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                
                If FiltroGeral = "Todos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codgrupo = b.codigo where a.ativo is null or a.ativo ='N'"
                'If FiltroGeral = "Todos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codcoligada = '" & vCodcoligada & "' and a.codgrupo = b.codigo and a.codcoligada = b.codcoligada"
                'If FiltroGeral = "Ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codcoligada = '" & vCodcoligada & "' and a.codgrupo = b.codigo and a.codcoligada = b.codcoligada where a.ativo = 'S'"
                'If FiltroGeral = "Não ativos" Then SqlLV = "select a.codigo,a.nome,b.descricao,a.ativo from tbusuarios as a inner join tbgrupo as b on a.codcoligada = '" & vCodcoligada & "' and a.codgrupo = b.codigo and a.codcoligada = b.codcoligada where a.ativo is null or a.ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Código", "Nome do usuário", "Grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Usuários"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 3, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'GRUPOS
        ElseIf QualLV = 14 Then
            Set chamaForm = New frmGrupos
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Grupos"
            LegendaExc = "Grupo" 'Usado na mensagem de exclusão
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select * from tbgrupo"
                If FiltroGeral = "Ativos" Then SqlLV = "select * from tbgrupo where ativo = 'S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select * from tbgrupo where ativo is null or ativo ='N'"
'                If FiltroGeral = "Todos" Then SqlLV = "select * from tbgrupo where codcoligada = '" & vCodcoligada & "'"
'                If FiltroGeral = "Ativos" Then SqlLV = "select * from tbgrupo where codcoligada = '" & vCodcoligada & "' and ativo = 'S'"
'                If FiltroGeral = "Não ativos" Then SqlLV = "select * from tbgrupo where codcoligada = '" & vCodcoligada & "' and ativo is null or codcoligada = '" & vCodcoligada & "' and ativo ='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Código", "Nome do grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Usuários"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'PROCESSO SELETIVO
        ElseIf QualLV = 15 Then
            Set chamaForm = New frmProcSel
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "Processo Seletivo"
            LegendaExc = "Processo Seletivo" 'Usado na mensagem de exclusão
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select codprocesso,codrequisição,datainicio,datafim,status,ativo from tbprocessos where codcoligada = '" & vCodcoligada & "'"
                If FiltroGeral = "Ativos" Then SqlLV = "select codprocesso,codrequisição,datainicio,datafim,status,ativo from tbprocessos where codcoligada = '" & vCodcoligada & "' and ativo='S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select codprocesso,codrequisição,datainicio,datafim,status,ativo from tbprocessos where codcoligada = '" & vCodcoligada & "' and ativo='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Requisição", "Data inicio", "Data fim", "Status", "Ativo", "", "", "", "", "", "", "", "", "", ""
            DimensionaLV "Processos Seletivos"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 1, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 5, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'INTD - Identificação das Necessidades de Treinamento e Desenvolvimento
        ElseIf QualLV = 16 Then
            Set chamaForm = New frmINTD
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "INTD"
            LegendaExc = "INTD" 'Usado na mensagem de exclusão
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,a.ativo from tbINTD as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador = b.id"
                If FiltroGeral = "Ativos" Then SqlLV = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,a.ativo from tbINTD as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador = b.id where a.ativo='S'"
                If FiltroGeral = "Não ativos" Then SqlLV = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,a.ativo from tbINTD as a inner join tbcolaboradores as b a.codcoligada = '" & vCodcoligada & "' and on a.codcolaborador = b.id where a.ativo='N'"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Data inicio", "Data Término", "Registro", "Colaborador", "Status", "Ativo", "", "", "", "", "", "", "", "", ""
            DimensionaLV "INTD - Identificação das Necessidades de Treinamento e Desenvolvimento"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                'PersonaColLV 0, "N", "N", "", "N", "S", "N", "D"
                PersonaColLV 6, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            'If Tipo = True Then MeuLV.Show 1
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'PDO - Processo Decisório Organizacional
        ElseIf QualLV = 17 Then
            Set chamaForm = New frmPDO
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "PDO"
            LegendaExc = "PDO" 'Usado na mensagem de exclusão
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "Select a.id,a.decisao,a.status,a.cpf,b.nomecolaborador,a.tipo,a.nota,a.solicitacao,a.datasolicitacao,a.solicitante,a.aprovador from tbAutorizacao as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf"
                If FiltroGeral = "Avaliados" Then SqlLV = "Select a.id,a.decisao,a.status,a.cpf,b.nomecolaborador,a.tipo,a.nota,a.solicitacao,a.datasolicitacao,a.solicitante,a.aprovador from tbAutorizacao as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf where a.status is not null"
                If FiltroGeral = "Não Avaliados" Then SqlLV = "Select a.id,a.decisao,a.status,a.cpf,b.nomecolaborador,a.tipo,a.nota,a.solicitacao,a.datasolicitacao,a.solicitante,a.aprovador from tbAutorizacao as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf where a.status is null"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            QtdColReal = 0
            MontaCabLV "Identificador", "Decisão", "Status", "CPF", "Nome", "Tipo", "Resultado", "Solicitação", "Data Sol.", "Solicitante", "Aprovado por", "", "", "", "", ""
            DimensionaLV "PDO - Processo Decisório Organizacional"
            MontaCabecalhoLV
            MontaDadosLV "S"
            If checaFiltro = True Then
                PersonaColLV 2, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 6, "S", "S", "%", "N", "N", "S", "D"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(5).Picture
            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(6).Picture
            MeuLV.cmdconsulta(4).ToolTipText = "Avaliar"
            MeuLV.cmdconsulta(5).ToolTipText = "Sair"
            MeuLV.cmdconsulta(6).Visible = False
            MeuLV.cmdconsulta(7).Visible = False
            'If Tipo = True Then MeuLV.Show
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        'ADP - Avaliação de Desempenho Profissional
        ElseIf QualLV = 18 Then
            Set chamaForm = New frmADP
            'frmMenu2.aicAlphaImage1.Visible = True
            Formulario = "ADP"
            LegendaExc = "ADP" 'Usado na mensagem de exclusão
            indiceVarGlobal = 1
            If Pesquisa <> "excluir" And Pesquisa <> "novo" And Pesquisa <> "editar" And Pesquisa <> "0" Then
                MontaFiltro
                If FiltroGeral = "" Then frmFiltro.Show 1
                If MeuLV.Visible = True Then Unload MeuLV
                If FiltroGeral = "Todos" Then SqlLV = "select b.codcolaborador,b.nomecolaborador,a.tipoadp,cast(a.dias as int),CONVERT (VARCHAR, a.datavencimento, 103) as datavencimento,CONVERT (VARCHAR, a.datadevolucao, 103) as datadevolucao,CONVERT (VARCHAR, a.dataavaliacao, 103) as dataavaliacao ,a.nota,a.statusimpressao,a.statusavaliacao,a.ativo,a.id from tbListaADP as a inner join tbColaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador=b.id and b.ativo = 'S' and b.tipo = 'colaborador' order by a.dias"
                If FiltroGeral = "Ativos" Then SqlLV = "select b.codcolaborador,b.nomecolaborador,a.tipoadp,cast(a.dias as int),CONVERT (VARCHAR, a.datavencimento, 103) as datavencimento,CONVERT (VARCHAR, a.datadevolucao, 103) as datadevolucao,CONVERT (VARCHAR, a.dataavaliacao, 103) as dataavaliacao ,a.nota,a.statusimpressao,a.statusavaliacao,a.ativo,a.id from tbListaADP as a inner join tbColaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador=b.id and b.ativo = 'S' and b.tipo = 'colaborador' where a.ativo is not null order by a.dias"
                If FiltroGeral = "Não ativos" Then SqlLV = "select b.codcolaborador,b.nomecolaborador,a.tipoadp,cast(a.dias as int),CONVERT (VARCHAR, a.datavencimento, 103) as datavencimento,CONVERT (VARCHAR, a.datadevolucao, 103) as datadevolucao,CONVERT (VARCHAR, a.dataavaliacao, 103) as dataavaliacao ,a.nota,a.statusimpressao,a.statusavaliacao,a.ativo,a.id from tbListaADP as a inner join tbColaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador=b.id and b.ativo = 'S' and b.tipo = 'colaborador' where a.ativo is null order by a.dias"
            Else
                If MeuLV.Visible = True Then Unload MeuLV
            End If
            If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
            MeuLV.cmdconsulta(10).Visible = True
            
            QtdColReal = 0
            MontaCabLV "Registro", "Nome", "Tipo", "Periodo", "Vencimento", "Devolução", "Avaliado em", "Pontuação", "Impresso", "Status ADP", "Ativo", "id", "", "", "", ""
            DimensionaLV "ADP - Avaliação de Desempenho Profissional"
            MontaCabecalhoLV
            MontaDadosLV "N"
            If checaFiltro = True Then
                PersonaColLV 3, "N", "N", "", "N", "N", "N", "D"
                PersonaColLV 7, "S", "S", "%", "N", "N", "S", "D"
                PersonaColLV 8, "N", "N", "", "S", "N", "N", "E"
                PersonaColLV 10, "N", "N", "", "S", "N", "N", "E"
            End If
            If MeuLV.ListView1.ListItems.Count > 0 Then ajusta_LV
            MeuLV.Label2.Caption = FiltroGeral
            CompoeComboLV MeuLV.Combo1
            Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(5).Picture
            Set MeuLV.cmdconsulta(5).PictureNormal = MeuLV.ImageList1.ListImages(6).Picture
            MeuLV.cmdconsulta(4).ToolTipText = "Avaliar"
            MeuLV.cmdconsulta(5).ToolTipText = "Sair"
            MeuLV.cmdconsulta(6).Visible = False
            MeuLV.cmdconsulta(7).Visible = False
            'If Tipo = True Then MeuLV.Show 1
            'frmMenu2.aicAlphaImage1.Visible = False
            Exit Sub
        End If
        Set frmFiltro = Nothing
        Set MeuLV = Nothing
        Set chamaForm = Nothing
End Sub

Public Sub CarregaSQLExcluir(QLV As Integer)
    Dim rsExcLVGeral As New ADODB.Recordset
    Dim P As Integer
    If QLV = 0 Then
        frmDemitirColaborador.Show 1
        gravaLog varGlobal, MeuLV.ListView1.SelectedItem.ListSubItems.Item(1), "-"
    ElseIf QLV = 1 Then
        SqlExcLVGeral = "Delete from tbColaboradores where a.codcoligada = '" & vCodcoligada & "' and cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresesc where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradorescur where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoresexp where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'; Delete from tbColaboradoreshist where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'"
    ElseIf QLV = 2 Then
        'SqlExcLVGeral = "Delete from tbDepartamentos where codDepartamento= '" & Val(varGlobal) & "' ;Delete from tbDepartamentosHistResp where codDepartamento= '" & Val(varGlobal) & "'"
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbDepartamentos set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and coddepartamento = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 3 Then
        'SqlExcLVGeral = "Delete from tbSetores where codSetor= '" & Val(varGlobal) & "' ;Delete from tbSetoresHistResp where codSetor= '" & Val(varGlobal) & "'"
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
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
        'NAO EXCLUI, DESATIVA OS CARGOS SELECIONADOS PARA EXCLUSÃO
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbcargos set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codcargo = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 5 Then
        'SqlExcLVGeral = "Delete from tbHabilidades where codHabilidade= " & Val(varGlobal)
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
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
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
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
        SqlExcLVGeral = "Delete from tbRequisicoes where codcoligada = '" & vCodcoligada & "' and codrequisicao= '" & Val(varGlobal) & "' ;Delete from tbRequisicoesAprovadores where codcoligada = '" & vCodcoligada & "' and codrequisicao= '" & Val(varGlobal) & "' ;Delete from tbRequisicoesCargos where codcocoligada = '" & vCodcoligada & "' and codrequisicao= '" & Val(varGlobal) & "'"
    ElseIf QLV = 8 Then
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            SqlExcLVGeral = "Select count(*) from tbmatrizcur as a where a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = '" & Val(varGlobal) & "'"
            rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
            If rsExcLVGeral.Fields(0) > 0 Then ' Desativa
                rsExcLVGeral.Close
                SqlExcLVGeral = "UPDATE tbTreinamentos set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(varGlobal) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                mobjMsg.Abrir "Curso/treinamento DESATIVADO com sucesso", Ok, informacao, "SGC"
            Else 'Verifica em outra tabela
                rsExcLVGeral.Close
                SqlExcLVGeral = "Select count(*) from tbcolaboradorescur as a where a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = '" & Val(varGlobal) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
                If rsExcLVGeral.Fields(0) > 0 Then ' Desativa
                    rsExcLVGeral.Close
                    SqlExcLVGeral = "UPDATE tbTreinamentos set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(varGlobal) & "'"
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                    mobjMsg.Abrir "Curso/treinamento DESATIVADO com sucesso", Ok, informacao, "SGC"
                Else 'Exclui
                    rsExcLVGeral.Close
                    SqlExcLVGeral = "Delete from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(varGlobal) & "'"
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                    mobjMsg.Abrir "Curso/treinamento EXCLUIDO com sucesso", Ok, informacao, "SGC"
                End If
            End If
        End If
        cnBanco.CommitTrans
        'rsExcLVGeral.Close
        'SqlExcLVGeral = "Delete from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codTreinamento=  '" & Val(varGlobal) & "' ;Delete from tbTreinamentosRev where codcoligada = '" & vCodcoligada & "' and codTreinamento= '" & Val(varGlobal) & "'"
    ElseIf QLV = 9 Then
        SqlExcLVGeral = "Select a.codmatriz from tbmatriz as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = b.codmatriz where a.codmatriz = '" & Val(varGlobal) & "'"
        rsExcLVGeral.Open SqlExcLVGeral, cnBanco, adOpenKeyset, adLockReadOnly
        If rsExcLVGeral.RecordCount = 0 Then
            rsExcLVGeral.Close
            Set rsExcLVGeral = Nothing
            mobjMsg.Abrir "Confirma exclusão da " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
            If Tp = 1 Then
                SqlExcLVGeral = "Delete from tbMatriz where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizCur where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizEsc where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizExp where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "' ;Delete from tbMatrizHab where codcoligada = '" & vCodcoligada & "' and codMatriz= '" & Val(varGlobal) & "'"
                rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                mobjMsg.Abrir "Registro excluido com sucesso", Ok, informacao, "SGC"
            End If
        Else
            rsExcLVGeral.Close
            Set rsExcLVGeral = Nothing
            mobjMsg.Abrir "Matriz não poder ser excluida! A Chave primária está sendo utilizada em outras tabelas", Ok, critico, "Atenção"
        End If
    ElseIf QLV = 10 Then
        Dim strResultado As String
        mobjMsg.Abrir "Confirma o Cancelamento da " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            While strResultado = ""
                strResultado = InputBox("Justifique o cancelamento da programação", "Cancelar Programação")
                If StrPtr(strResultado) = 0 Then
                    mobjMsg.Abrir "Cancelamento de programação foi cancelado", Ok, informacao, "Atenção"
                    Exit Sub
                End If
                If strResultado = "" Then
                    mobjMsg.Abrir "É necessário justificar o cancelamento", Ok, critico, "Atenção"
                End If
            Wend
            
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    'If strResultado <> "" Then
                        SqlExcLVGeral = "UPDATE tbPendentesCur set ativo = 'N', status = 'Cancelado', observacao = '" & strResultado & "' where codcoligada = '" & vCodcoligada & "' and cpf = '" & MeuLV.ListView1.ListItems.Item(P) & "' and codtreinamento = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)) & "' and status <> 'Concluido'"
                        rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                    'Else
                    '    MsgBox "É necessário justificar o cancelamento"
                    'End If
                End If
            Next
            mobjMsg.Abrir "Cancelamento realizado!", Ok, critico, "Atenção"
        End If
    ElseIf QLV = 11 Then
        'SqlExcLVGeral = "Delete from tbAvaliacao where codavaliacao= " & Val(varGlobal)
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            For P = 1 To MeuLV.ListView1.ListItems.Count
                If MeuLV.ListView1.ListItems.Item(P).Checked = True Then
                    SqlExcLVGeral = "UPDATE tbAvaliacao set ativo = 'N' where codcoligada = '" & vCodcoligada & "' and codavaliacao = " & Val(MeuLV.ListView1.ListItems.Item(P))
                    rsExcLVGeral.Open SqlExcLVGeral, cnBanco
                End If
            Next
        End If
        cnBanco.CommitTrans
    ElseIf QLV = 16 Then
        cnBanco.BeginTrans
        mobjMsg.Abrir "Confirma exclusão do " & LegendaExc & " selecionado?", YesNo, pergunta, "SGC"
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

'Gera Avaliação de Desempenho Profissional por colaborador
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
End Function
