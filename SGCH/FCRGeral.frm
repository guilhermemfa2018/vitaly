VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRGeral 
   Caption         =   "Listagem"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRGeral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      _cx             =   10231
      _cy             =   12347
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1046
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "FCRGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRGeral

Private Sub Form_Load()
    Dim report1 As New CRGeral
    Dim CRXApplication As New CRAXDDRT.Application
    Set CRXApplication = CreateObject("CrystalRuntime.Application.11")
    Dim CRXReport As New CRAXDDRT.Report
    Dim CRXDatabase As CRAXDDRT.Database
    Set CRXReport = report1
    Set CRXDatabase = CRXReport.Database
    
    Dim rsGeral As New ADODB.Recordset
    Dim SqlGeral As String
    
    rsGeral.CursorLocation = adUseClient
    If apontaLV = 2 Then
        'DEPARTAMENTO
        SqlGeral = "select a.coddepartamento as campo1,a.nomedepartamento as campo2,a.descricao as campo3,'-' as campo4,'CÓDIGO' as campo5,'DEPARTAMENTO' as campo6,'DESCRIÇÃO' as campo7,'-' as campo8,'DEPARTAMENTOS' as campo9,c.logo as campo10 from tbDepartamentos as a inner join tbDadosEmpresa as c on c.codcoligada = '" & vCodcoligada & "' where a.ativo = 'S' order by a.nomedepartamento"
    ElseIf apontaLV = 0 Then
        'COLABORADORES
        'SqlGeral = "select a.coddepartamento as campo1,'-' as campo2,'-' as campo3,'-' as campo4,'' as campo5,'' as campo6,'' as campo7,'-' as campo8,'' as campo9,'-' as campo10 from tbDepartamentos as a where a.ativo = 'S' order by a.nomedepartamento"
        'SqlGeral = "select a.id as campo1,'-' as campo2,'-' as campo3, '-' as campo4,'-' as campo5,'-' as campo6,'-' as campo7,'-' as campo8,'-' as campo9,'-' as campo10 from tbcolaboradores as a order by a.nomecolaborador"
    
        SqlGeral = "select a.id AS campo1,substring(a.nomecolaborador,1,20) + ' - ' +CONVERT (VARCHAR, b.data, 103) as campo2,d.nomecargo as campo3,cast(cast(a.mediageral as decimal(7,2)) as varchar) + '% - ' + Max(f.nomeescolaridade)  as campo4," & _
                    "'ID' as campo5,'ADMISSÃO' as campo6,'CARGO' as campo7,'PONTUAÇÃO/ESCOLARIDADE' as campo8,'ESCOLARIDADE' as campo9,g.logo as campo10 from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and " & _
                    "a.cpf = b.cpf and a.ativo = 'S' and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo left join tbColaboradoresEsc as e on a.cpf = e.cpf " & _
                    "left join tbescolaridade as f on e.codescolaridade = f.codescolaridade inner join tbDadosEmpresa as g on g.codcoligada = '" & vCodcoligada & "' group by a.id,a.nomecolaborador,b.data,d.nomecargo,a.mediageral,g.logo order by a.id"
    
    ElseIf apontaLV = 3 Then
        'SETOR
        SqlGeral = "select a.codsetor as campo1,a.nomesetor as campo2,b.nomedepartamento as campo3,a.descricao as campo4,'CÓDIGO' as campo5,'SETOR' as campo6,'DEPARTAMENTO' as campo7,'DESCRIÇÃO' as campo8,'SETORES' as campo9,c.logo as campo10 from tbSetores as a inner join tbdepartamentos as b on a.coddepartamento = b.coddepartamento inner join tbDadosEmpresa as c on c.codcoligada = '" & vCodcoligada & "' where a.ativo = 'S'  order by a.nomesetor"
    ElseIf apontaLV = 6 Then
        'ESCOLARIDADE
        SqlGeral = "select a.codescolaridade as campo1,a.nomeescolaridade as campo2,CAST(a.peso AS VARCHAR) as campo3,'-' as campo4,'CÓDIGO' as campo5,'ESCOLARIDADE' as campo6,'PESO' as campo7,'-' as campo8,'ESCOLARIDADES' as campo9,b.logo as campo10 from tbEscolaridade as a inner join tbDadosEmpresa as b on b.codcoligada = '" & vCodcoligada & "' where a.ativo = 'S' order by a.nomeescolaridade"
    ElseIf apontaLV = 5 Then
        'HABILIDADE
        SqlGeral = "select a.codhabilidade as campo1,a.nomehabilidade as campo2,CAST(a.peso AS VARCHAR) as campo3,a.descricao as campo4,'CÓDIGO' as campo5,'ESCOLARIDADE' as campo6,'PESO' as campo7,'DESCRIÇÃO' as campo8,'HABILIDADES' as campo9,b.logo as campo10 from tbHabilidades as a inner join tbDadosEmpresa as b on b.codcoligada = '" & vCodcoligada & "' where a.ativo = 'S' order by a.nomehabilidade"
    ElseIf apontaLV = 11 Then
        'AVALIACAO
        SqlGeral = "select a.codavaliacao as campo1,a.nomeavaliacao as campo2,a.tipo as campo3,a.descricao as campo4,'CÓDIGO' as campo5,'AVALIAÇÃO' as campo6,'TIPO'as campo7,'DESCRIÇÃO' as campo8,'AVALIAÇÕES' as campo9,b.logo as campo10 from tbAvaliacao as a inner join tbDadosEmpresa as b on b.codcoligada = '" & vCodcoligada & "' where a.ativo = 'S' order by a.codavaliacao"
    ElseIf apontaLV = 17 Then
        'PDO
        SqlGeral = "select a.id as campo1, substring(a.aprovador+space(25),1,25) +'  '+ a.decisao as campo2,b.nomecolaborador as campo3,a.nota  + '% - ' + a.solicitacao as campo4,'ID' as campo5,'APROVADOR              DECISÃO' as campo6,'COLABORADOR' as campo7,'NOTA' as campo8,'PROCESSO DECISÓRIO ORGANIZACIONAL' AS campo9,c.logo as campo10 " & _
                   "from tbautorizacao as a inner join tbcolaboradores as b on b.cpf = a.cpf and a.codcoligada = '" & vCodcoligada & "', tbdadosempresa as c order by a.id"
    End If
    
    
    rsGeral.Open SqlGeral, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsGeral.ActiveConnection = Nothing
    CRXReport.DiscardSavedData
    CRXReport.Database.SetDataSource rsGeral
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = CRXReport
    CRViewer1.ViewReport
    CRViewer1.Zoom (100)
    Screen.MousePointer = vbDefault


'Screen.MousePointer = vbHourglass
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRGeral.Hide
    Unload Me
    Set rsGeral = Nothing
End Sub

