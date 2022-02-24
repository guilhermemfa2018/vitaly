VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRAvaHab 
   Caption         =   "Avaliação de habilidades funcionais"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRAvaHab.frx":0000
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
      DisplayGroupTree=   -1  'True
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
Attribute VB_Name = "FCRAvaHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRAvaHab

Private Sub Form_Load()
    
    Dim report1 As New CRAvaHab
    Dim rsAvaHab As New ADODB.Recordset
    Dim SqlAvaHab As String
    
    rsAvaHab.CursorLocation = adUseClient
    If FiltroGeral = "Ativos" Then
        SqlAvaHab = "select a.cpf,a.codcolaborador,a.nomecolaborador,a.foto,c.nomehabilidade,c.descricao,b.pontuacao,f.nomecargo,g.logo,i.mediaaprovacao,i.aprovadorest,j.nomedepartamento,k.nomesetor from " & _
                   "tbcolaboradores as a inner join tbcolaboradoreshab as b on a.cpf = b.cpf and a.tipo = 'colaborador' inner join tbhabilidades as c on b.codhabilidade = c.codhabilidade inner join tbcolaboradoreshist as d on a.cpf = d.cpf and " & _
                   "d.ativo = 'S' and d.codmatriz = b.codmatriz inner join tbmatriz as e on d.codmatriz = e.codmatriz inner join tbdepartamentos as j on e.coddepartamento = j.coddepartamento inner join tbsetores as k on e.codsetor = k.codsetor " & _
                   "inner join tbcargos as f on e.codcargo = f.codcargo inner join tbDadosEmpresa as g on g.codcoligada = '" & vCodColigada & "', tbParametros as i where a.ativo = 'S'"
    ElseIf FiltroGeral = "Não ativos" Or FiltroGeral = "Demitidos" Then
        SqlAvaHab = "select a.cpf,a.codcolaborador,a.nomecolaborador,a.foto,c.nomehabilidade,c.descricao,b.pontuacao,f.nomecargo,g.logo,i.mediaaprovacao,i.aprovadorest,j.nomedepartamento,k.nomesetor from " & _
                   "tbcolaboradores as a inner join tbcolaboradoreshab as b on a.cpf = b.cpf and a.tipo = 'colaborador' inner join tbhabilidades as c on b.codhabilidade = c.codhabilidade inner join tbcolaboradoreshist as d on a.cpf = d.cpf and " & _
                   "d.ativo = 'S' and d.codmatriz = b.codmatriz inner join tbmatriz as e on d.codmatriz = e.codmatriz inner join tbdepartamentos as j on e.coddepartamento = j.coddepartamento inner join tbsetores as k on e.codsetor = k.codsetor " & _
                   "inner join tbcargos as f on e.codcargo = f.codcargo inner join tbDadosEmpresa as g on g.codcoligada = '" & vCodColigada & "', tbParametros as i where a.ativo = 'N'"
    ElseIf FiltroGeral = "Todos" Then
        SqlAvaHab = "select a.cpf,a.codcolaborador,a.nomecolaborador,a.foto,c.nomehabilidade,c.descricao,b.pontuacao,f.nomecargo,g.logo,i.mediaaprovacao,i.aprovadorest,j.nomedepartamento,k.nomesetor from " & _
                   "tbcolaboradores as a inner join tbcolaboradoreshab as b on a.cpf = b.cpf and a.tipo = 'colaborador' inner join tbhabilidades as c on b.codhabilidade = c.codhabilidade inner join tbcolaboradoreshist as d on a.cpf = d.cpf and " & _
                   "d.ativo = 'S' and d.codmatriz = b.codmatriz inner join tbmatriz as e on d.codmatriz = e.codmatriz inner join tbdepartamentos as j on e.coddepartamento = j.coddepartamento inner join tbsetores as k on e.codsetor = k.codsetor " & _
                   "inner join tbcargos as f on e.codcargo = f.codcargo inner join tbDadosEmpresa as g on g.codcoligada = '" & vCodColigada & "', tbParametros as i"
    End If
    rsAvaHab.Open SqlAvaHab, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsAvaHab.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsAvaHab
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (90)
    Screen.MousePointer = vbDefault
    
    rsAvaHab.Close
    Set rsAvaHab = Nothing
    Exit Sub
    
    
'    Screen.MousePointer = vbHourglass
'    CRViewer1.ReportSource = Report
'    CRViewer1.ViewReport
'    Screen.MousePointer = vbDefault
'    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRAvaHab.Hide
    Unload Me
    Set FCRAvaHab = Nothing
End Sub
