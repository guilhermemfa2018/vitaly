VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form FCRIntdGeral 
   Caption         =   "Relatório Geral de INTD's"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRIntdGeral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
Attribute VB_Name = "FCRIntdGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRIntdGeral

Private Sub Form_Load()

    Dim report1 As New CRIntdGeral
    Dim rsIntdGeral  As New ADODB.Recordset
    Dim SqlIntdGeral As String
    
    rsIntdGeral.CursorLocation = adUseClient
    
    'Cancelados ou Demitidos
    If Pesquisa = "Cancelada" Then
        SqlIntdGeral = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,e.nomecargo as [CARGO ATUAL],g.nomecargo as [CARGO TREINAMENTO],case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,h.logo " & _
                       "from tbintd as a inner join tbColaboradores as b on b.id = a.codcolaborador inner join tbColaboradoresHist as c on b.cpf = c.cpf and c.ativo = 'S' inner join tbMatriz as d on c.codmatriz = d.codmatriz inner join tbCargos as e on d.codcargo = e.codcargo " & _
                       "inner join tbMatriz as f on a.codmatriz = f.codmatriz inner join tbCargos as g on f.codcargo = g.codcargo,tbDadosEmpresa as h where status like '%" & Pesquisa & "%' or b.datarecisao is not null order by b.nomecolaborador"
                   
    'Em Andamento ou Aberto
    ElseIf Pesquisa = "Em andamento" Or Pesquisa = "Aberto" Then
        SqlIntdGeral = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,e.nomecargo as [CARGO ATUAL],g.nomecargo as [CARGO TREINAMENTO],case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,h.logo " & _
                       "from tbintd as a inner join tbColaboradores as b on b.id = a.codcolaborador inner join tbColaboradoresHist as c on b.cpf = c.cpf and c.ativo = 'S' inner join tbMatriz as d on c.codmatriz = d.codmatriz inner join tbCargos as e on d.codcargo = e.codcargo " & _
                       "inner join tbMatriz as f on a.codmatriz = f.codmatriz inner join tbCargos as g on f.codcargo = g.codcargo,tbDadosEmpresa as h where status like '%Aberto%' and  b.datarecisao is null or status like '%Em andamento%' and  b.datarecisao is null order by b.nomecolaborador"
    'Fechado
    ElseIf Pesquisa = "Fechado" Then
        SqlIntdGeral = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,e.nomecargo as [CARGO ATUAL],g.nomecargo as [CARGO TREINAMENTO],case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,h.logo " & _
                       "from tbintd as a inner join tbColaboradores as b on b.id = a.codcolaborador inner join tbColaboradoresHist as c on b.cpf = c.cpf and c.ativo = 'S' inner join tbMatriz as d on c.codmatriz = d.codmatriz inner join tbCargos as e on d.codcargo = e.codcargo " & _
                       "inner join tbMatriz as f on a.codmatriz = f.codmatriz inner join tbCargos as g on f.codcargo = g.codcargo,tbDadosEmpresa as h where a.status like '%" & Pesquisa & "%' order by b.nomecolaborador"
    'Todos
    ElseIf Pesquisa = "" Then
        SqlIntdGeral = "select a.codINTD,a.datainicio,a.datafim,b.codcolaborador,b.nomecolaborador,e.nomecargo as [CARGO ATUAL],g.nomecargo as [CARGO TREINAMENTO],case when b.datarecisao is not null then 'Demitido' when b.datarecisao is null then a.status end as Status,h.logo " & _
                       "from tbintd as a inner join tbColaboradores as b on b.id = a.codcolaborador inner join tbColaboradoresHist as c on b.cpf = c.cpf and c.ativo = 'S' inner join tbMatriz as d on c.codmatriz = d.codmatriz inner join tbCargos as e on d.codcargo = e.codcargo " & _
                       "inner join tbMatriz as f on a.codmatriz = f.codmatriz inner join tbCargos as g on f.codcargo = g.codcargo,tbDadosEmpresa as h order by b.nomecolaborador"
    End If
    rsIntdGeral.Open SqlIntdGeral, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsIntdGeral.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsIntdGeral
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (102)
    Screen.MousePointer = vbDefault
    
    rsIntdGeral.Close
    Set rsIntdGeral = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRIntdGeral.Hide
    Unload Me
    Set FCRIntdGeral = Nothing
End Sub


