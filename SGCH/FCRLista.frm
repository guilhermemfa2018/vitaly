VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRLista 
   Caption         =   "Relatorio Geral - Colaboradores"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRLista.frx":0000
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
Attribute VB_Name = "FCRLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRLista

Private Sub Form_Load()
    Dim report1 As New CRLista
    Dim rsRelLista As New ADODB.Recordset
    Dim sqlRelLista As String
    
    rsRelLista.CursorLocation = adUseClient
    sqlRelLista = "select CASE WHEN a.ativo = 'S' THEN 'Ativo' WHEN a.ativo = 'N' THEN 'Demitido' ELSE '-' END STATUS,a.id AS ID,substring(a.nomecolaborador,1,50) as NOME,a.cpf as CPF,a.codcolaborador as REGISTRO, " & _
                  "CASE WHEN a.ctpsnumero = '' THEN '-' ELSE a.ctpsnumero + ' - Serie:' + a.ctpsserie END as CTPS,d.nomecargo as CARGO,CONVERT (VARCHAR, b.data, 103) as ADMISSAO,CONVERT (VARCHAR, a.datanascimento, 103) as NASCIMENTO, " & _
                  "cast(cast(a.mediageral as decimal(7,2)) as varchar) as PONTU, Max(f.nomeescolaridade) as ESCOLARIDADE,g.logo as LOGO from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and " & _
                  "a.cpf = b.cpf inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo left join tbColaboradoresEsc as e on a.cpf = e.cpf left join tbescolaridade as f on " & _
                  "e.codescolaridade = f.codescolaridade inner join tbDadosEmpresa as g on g.codcoligada = '" & vCodcoligada & "' group by a.ativo,a.id,a.nomecolaborador,a.cpf,a.codcolaborador,a.ctpsnumero,a.ctpsserie,b.data,a.datanascimento,d.nomecargo,a.mediageral,g.logo order by a.ativo desc,a.nomecolaborador"

    
    '"select a.codprogramacao,f.cpf,a.datainicio,a.horainicio,a.local,e.cargahoraria,b.nomeinstrutor,e.nometreinamento,c.codavaliacao,c.nomeavaliacao,c.descricao,d.logo,i.mediaaprovacao,i.aprovadorest from " & _
    '             "tbprogramacao as a inner join tbpendentescur as f on a.codprogramacao = f.codprogramacao inner join tbtreinamentos as e on e.codtreinamento = f.codtreinamento inner join tbProgramacaoInstrutores as b " & _
    '             "on a.codprogramacao = b.codprogramacao,tbavaliacao as c,tbDadosEmpresa as d inner join tbparametros as i on i.codcoligada = '" & vCodcoligada & "' where c.tipo = 'AT' and c.ativo = 'S'"
    
    rsRelLista.Open sqlRelLista, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsRelLista.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsRelLista
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsRelLista.Close
    Set rsRelLista = Nothing
    
    
    
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRLista.Hide
    Unload Me
    Set FCRLista = Nothing
End Sub


