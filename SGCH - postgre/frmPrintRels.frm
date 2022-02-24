VERSION 5.00
Begin VB.Form frmPrintRels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmPrintRels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selecione o relatório que deseja visualizar "
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmPrintRels.frx":3469A
         Left            =   120
         List            =   "frmPrintRels.frx":346AD
         TabIndex        =   1
         Text            =   "Histórico funcional"
         Top             =   360
         Width           =   4215
      End
   End
   Begin SGCH.chameleonButton cmdImprimir 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintRels.frx":34728
      PICN            =   "frmPrintRels.frx":34744
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdImprimir 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintRels.frx":3541E
      PICN            =   "frmPrintRels.frx":3543A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmPrintRels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click(Index As Integer)
On Error Resume Next
    Select Case Index
    Case 0
        If apontaLV = 0 Then
            If Combo1.ListIndex = 0 Then
                montaTbPrintHFunc
                FCRHistFunc.Show 1
            ElseIf Combo1.ListIndex = 1 Then
                FCRTreiInt.Show 1
            ElseIf Combo1.ListIndex = 2 Then
                FCRAvaHab.Show 1
            ElseIf Combo1.ListIndex = 3 Then
                FCRListaColCargos.Show 1
            ElseIf Combo1.ListIndex = 4 Then
                FCRGeral.Show 1
            End If
        ElseIf apontaLV = 18 Then
            If Combo1.ListIndex = 0 Then
                If MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = "-" Then
                    MsgBox "ADP não avaliada, não pode ser impressa", vbCritical, "SGCH"
                    Exit Sub
                End If
                If MeuLV.ListView1.ListItems.Count > 0 Then
                    varGlobal2 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(11)
                    FCRADP.Show 1
                Else
                    MsgBox "Nenhuma Avaliação de Desempenho foi selecionada", vbCritical, "SGCH"
                End If
            End If
        ElseIf apontaLV = 16 Then
            If Combo1.ListIndex = 0 Then
                If MeuLV.ListView1.ListItems.Count > 0 Then
                    SelectINTD
                    montaTbPrintINTD
                    FCRIntdIndividual.Show 1
                Else
                    MsgBox "Nenhuma INTD foi selecionada", vbCritical, "SGCH"
                End If
            ElseIf Combo1.ListIndex = 1 Then
                MsgComboBox "Selecione uma das opções abaixo", 0, 12, 0, 1
                If TiPo = 1 Then FCRIntdGeral.Show 1
            End If
        ElseIf apontaLV = 10 Then
            If Combo1.ListIndex = 0 Then
                frmConvocacao.Show 1
            ElseIf Combo1.ListIndex = 1 Then
                'Cria tabela temporária
                criaTabTempProg
                'Joga os dados selecionados no listview para a tabela temporaria
                insereDadosTemp
                'Criar o relatorio com os dados da tabela temporária
                FCRListaProg.Show 1
            End If
        End If
    Case 1
        Unload Me
        Set frmPrintRels = Nothing
    End Select
End Sub

Private Sub Form_Load()
    If apontaLV = 18 Then
        Combo1.Clear
        Combo1.AddItem "Avaliação de Desempenho Profissional"
    ElseIf apontaLV = 16 Then
        Combo1.Clear
        Combo1.AddItem "INTD Individual"
        Combo1.AddItem "INTD Geral"
    ElseIf apontaLV = 10 Then
        Combo1.Clear
        Combo1.AddItem "Convocação de treinamento"
        Combo1.AddItem "Programações gerais"
    End If
    Combo1.ListIndex = 0
End Sub

Private Sub montaTbPrintHFunc()
    Dim rsColaborador As New ADODB.Recordset
    Dim SqlColaborador As String
    
    Dim rsPrintMatriz As New ADODB.Recordset
    Dim SqlPrintMatriz As String
    Dim vCPFColaborador As String
    
    cnBanco.BeginTrans
    
    SqlPrintMatriz = "Delete from tbPrintHFunc where codcoligada = '" & vCodcoligada & "'"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
    SqlColaborador = "select * from tbColaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' order by cpf"
    rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsColaborador.EOF
        vCPFColaborador = rsColaborador.Fields(0)

        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "FUNÇÕES" na tabela tbPrintMatriz
'        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5) Values(" & vCPFColaborador & ",'FUNÇÕES' + space(89) + 'Período')"
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5,codcoligada) Values(REPLICATE('0', 11 - Len(" & vCPFColaborador & ")) + RTrim(" & vCPFColaborador & "),'FUNÇÕES' + space(89) + 'Período','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "FUNÇÕES" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,codcoligada) " & _
                         "Select a.cpf,d.nomecargo + ' - ' + c.nivel , substring(convert(char,a.data,103),1,10) + ' - ' + isnull(substring(convert(char,a.datasai,103),1,10),''),'" & vCodcoligada & "' " & _
                         "from tbColaboradoresHist as a inner join tbColaboradores as b on a.cpf = b.cpf inner join tbMatriz as c on a.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo " & _
                         "where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & vCPFColaborador & "' order by a.data desc"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
    
        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "HABILIDADES" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5,codcoligada) Values(REPLICATE('0', 11 - Len(" & vCPFColaborador & ")) + RTrim(" & vCPFColaborador & "),'HABILIDADES','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "HABILIDADES" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,codcoligada) Select a.cpf,b.nomehabilidade,'" & vCodcoligada & "' from tbColaboradoresHab as a inner join tbhabilidades as b on  a.codhabilidade = b.codhabilidade inner join tbcolaboradoreshist as c on  a.cpf = c.cpf and c.ativo = 'S' " & _
                         "where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & vCPFColaborador & "' and a.pontuacao >= '" & MediaGlobal & "' and c.codmatriz = a.codmatriz " & _
                         "group by a.cpf, a.codhabilidade, b.nomehabilidade"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "ESCOLARIDADE" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5,codcoligada) Values(REPLICATE('0', 11 - Len(" & vCPFColaborador & ")) + RTrim(" & vCPFColaborador & "),'ESCOLARIDADE','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "ESCOLARIDADE" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,codcoligada) Select a.cpf,b.nomeescolaridade,'" & vCodcoligada & "' from tbColaboradoresEsc as a inner join tbEscolaridade as b on a.codescolaridade = b.codescolaridade where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & vCPFColaborador & "'"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "CURSOS/TREINAMENTOS" na tabela tbPrintMatriz
        'SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5) Values(" & vCPFColaborador & ",'CURSOS/TREINAMENTOS')"
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5,codcoligada) Values(REPLICATE('0', 11 - Len(" & vCPFColaborador & ")) + RTrim(" & vCPFColaborador & "),'CURSOS/TREINAMENTOS' + space(61) + 'C.H.'+ space(16) + 'Período' + space(24) + 'Programação','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
        
        ''ABAIXO - Insere a "CURSOS/TREINAMENTOS" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4,codcoligada) Select a.cpf,b.nometreinamento + ' (' + isnull(c.nomenivel,'') + ')',substring(b.cargahoraria,1,3) + ':'+ substring(b.cargahoraria,4,2),substring(convert(char,e.datainicio,103),1,10) + ' - ' + substring(convert(char,e.datafim,103),1,10) + '       ' + REPLICATE ('0',6 - len(d.codprogramacao))+convert(char,d.codprogramacao),'" & vCodcoligada & "' " & _
                         "From tbcolaboradorescur as a Inner join tbtreinamentos as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento left join tbtreinamentosniv as c on a.codtreinamento = c.codtreinamento left join tbpendentescur as d " & _
                         "on d.situacao = 'Aprovado' and d.cpf = a.cpf and d.codtreinamento = a.codtreinamento and d.codprogramacao is not null left join tbprogramacao as e on d.codprogramacao = e.codprogramacao where a.cpf = '" & vCPFColaborador & "'" & _
                         " Group by a.cpf, b.nometreinamento + ' (' + isnull(c.nomenivel,'') + ')',substring(b.cargahoraria,1,3) + ':'+ substring(b.cargahoraria,4,2),substring(convert(char,e.datainicio,103),1,10) + ' - ' + substring(convert(char,e.datafim,103),1,10) + '       ' + REPLICATE ('0',6 - len(d.codprogramacao))+convert(char,d.codprogramacao)"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

'        ''ABAIXO - Insere a "CURSOS/TREINAMENTOS" referente a MATRIZ selecionada na tabela tbPrintMatriz
'        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4) Select a.cpf,c.nometreinamento + ' (' + isnull(g.nomenivel,'') + ')',substring(c.cargahoraria,1,3) + ':'+ substring(c.cargahoraria,4,2), substring(convert(char,e.datainicio,103),1,10) + ' - ' + substring(convert(char,e.datafim,103),1,10) + '       ' + REPLICATE ('0',6 - len(d.codtreinamento))+convert(char,d.codtreinamento) " & _
'                         "from tbColaboradoresCur as a inner join tbcolaboradores as b on a.cpf = b.cpf inner join tbtreinamentos as c on a.codtreinamento = c.codtreinamento left join tbpendentescur as d " & _
'                         "on a.cpf = d.cpf and c.codtreinamento = d.codtreinamento left join tbprogramacao as e on d.codprogramacao = e.codprogramacao left join tbprogramacaoinstrutores as f on e.codprogramacao = f.codprogramacao left join tbTreinamentosNiv as g on a.codtreinamento = g.codtreinamento and a.codnivel = g.codnivel " & _
'                         "where a.cpf = '" & vCPFColaborador & "'"
'        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
'        'Mid({Command.cargahoraria},1,3) & ":" & Mid({Command.cargahoraria},4,2)

        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "EXPERIÊNCIA" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo5,codcoligada) Values(REPLICATE('0', 11 - Len(" & vCPFColaborador & ")) + RTrim(" & vCPFColaborador & "),'EXPERIÊNCIA CURRICULAR' + space(57) + 'Tempo'+ space(12) + 'Empresa','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "EXPERIÊNCIA" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4,codcoligada) Select a.cpf,b.nomecargo,a.tempoexp,a.nomeempresa,'" & vCodcoligada & "' from tbColaboradoresExp as a inner join tbcargos as b on a.codcoligada = '" & vCodcoligada & "' and a.codcargo = b.codcargo where a.cpf = '" & vCPFColaborador & "'"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
        '******************************************************************************
        rsColaborador.MoveNext
    Wend
    cnBanco.CommitTrans
End Sub

Private Sub montaTbPrintINTD()
    Dim rsColaborador As New ADODB.Recordset
    Dim SqlColaborador As String
    
    Dim rsPrintMatriz As New ADODB.Recordset
    Dim SqlPrintMatriz As String
    Dim vCPFColaborador As String
    
    
    SqlPrintMatriz = "Delete from tbPrintHFunc where codcoligada = '" & vCodcoligada & "'"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
    '******************************************************************************
    '*** ABAIXO - Insere o cabeçalho de "CURSOS/TREINAMENTOS" na tabela tbPrintMatriz
    SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4,campo5,codcoligada) Values('CÓDIGO','NOME','PONTUAÇÃO','PROGRAMAÇÃO','CURSOS/TREINAMENTOS','" & vCodcoligada & "')"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco

    ''ABAIXO - Insere a "CURSOS/TREINAMENTOS" referente a MATRIZ selecionada na tabela tbPrintMatriz
    SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4,campo5,codcoligada) " & _
                     "Select a.codTreinamento, b.nometreinamento, d.nota, d.codprogramacao,'CURSOS/TREINAMENTOS','" & vCodcoligada & "' " & _
                     "from tbINTDcur as a left join tbTreinamentos as b on a.codtreinamento=b.codtreinamento " & _
                     "left join tbTreinamentosNiv as c on b.codtreinamento = c.codtreinamento and a.codnivel = c.codnivel " & _
                     "left join tbPendentesCur as d on a.codINTD = d.codINTD and a.codTreinamento = d.codtreinamento and d.codINTD = '" & Val(varGlobal2) & "' " & _
                     "where a.codcoligada = '" & vCodcoligada & "' and a.codINTD = '" & Val(varGlobal2) & "'"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
    
    '******************************************************************************
    '*** ABAIXO - Insere o cabeçalho de "HABILIDADES" na tabela tbPrintMatriz
    SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4,campo5,codcoligada) Values('CÓDIGO','NOME','PONTUAÇÃO','','HABILIDADES','" & vCodcoligada & "')"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco

    ''ABAIXO - Insere as "HABILIDADES" referente a INTD selecionada na tabela tbPrintMatriz
    SqlPrintMatriz = "Insert into tbPrintHFunc(campo1,campo2,campo3,campo4,campo5,codcoligada) " & _
                     "Select a.codHabilidade,b.nomehabilidade,a.pontuacao,'','HABILIDADES','" & vCodcoligada & "' " & _
                     "from tbINTDHab as a inner join tbHabilidades as b on a.codHabilidade = b.codhabilidade and a.codcoligada = '" & vCodcoligada & "' " & _
                     "where a.codcoligada = '" & vCodcoligada & "' and a.codINTD = '" & Val(varGlobal2) & "'"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco
End Sub

Private Sub criaTabTempProg()
On Error Resume Next
    'Criando uma tabela para impressão de programações
    Dim rsTabTempProg As New ADODB.Recordset
    Dim SqlTabTempProg As String
    SqlTabTempProg = "CREATE TABLE ##TempProg(CPF VARCHAR(50) NOT NULL,nomecolaborador VARCHAR(100) NOT NULL,treinamento VARCHAR(100) NOT NULL, funcaocolaborador VARCHAR(100) NOT NULL)"
    rsTabTempProg.Open SqlTabTempProg, cnBanco
End Sub

Private Sub insereDadosTemp()
    Dim rsGravaColaboradores As New ADODB.Recordset
    Dim sqlGravaColaboradores As String
    
    Dim Y As Integer, X As Integer
    
    Dim rsDeletaTemp As New ADODB.Recordset
    Dim sqlDeletaTemp As String
    
    sqlDeletaTemp = "delete from ##TempProg"
    rsDeletaTemp.Open sqlDeletaTemp, cnBanco
    
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        MeuLV.ListView1.ListItems.Item(X).Selected = True
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            sqlGravaColaboradores = "INSERT INTO ##TempProg(cpf,nomecolaborador,treinamento,funcaocolaborador) VALUES('" & MeuLV.ListView1.ListItems.Item(X) & "','" & MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) & "','" & MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) & "','" & MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) & "')"
            rsGravaColaboradores.Open sqlGravaColaboradores, cnBanco
        End If
    Next
End Sub

Private Sub SelectINTD()
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    varGlobal2 = MeuLV.ListView1.ListItems.Item(X)
End Sub
