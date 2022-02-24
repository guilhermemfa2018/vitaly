VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form FCRListaPresenca 
   Caption         =   "Relatorio de Programação"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRListaPresenca.frx":0000
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
Attribute VB_Name = "FCRListaPresenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRListaPresenca

Private Sub Form_Load()
    Dim report1 As New CRListaPresenca
    Dim rsRelProg As New ADODB.Recordset
    Dim sqlRelInsp As String
    
    rsRelProg.CursorLocation = adUseClient
    sqlRelInsp = "select a.codprogramacao,a.dataprogramacao,a.datainicio,a.datafim,a.horainicio,a.horafim,a.local,a.entidade,a.dae,a.metodo,a.metodooutro,a.observacao,a.avalnome,a.avaldata,g.nomeinstrutor as nomeinstrutor1,k.nomeinstrutor as nomeinstrutor2,c.codcolaborador,c.nomecolaborador,f.nomecargo,j.nomesetor,h.nometreinamento,h.objetivo,h.conteudo,h.cargahoraria,i.logo,h.origem,h.tipo,a.metodoA,a.metodoT,a.metodoS,a.metodoPT from " & _
                 "tbprogramacao as a inner join tbpendentescur as b on a.codprogramacao=b.codprogramacao inner join tbcolaboradores as c on b.cpf = c.cpf and c.ativo = 'S' inner join tbColaboradoresHist as d on c.cpf = d.cpf and d.ativo = 'S' inner join tbmatriz as e on e.codmatriz = d.codmatriz inner join tbcargos as f on f.codcargo = e.codcargo inner join tbsetores as j on j.codsetor = e.codsetor left join tbprogramacaoinstrutores as g on g.codprogramacao = a.codprogramacao and g.sequencia = 1 left join tbprogramacaoinstrutores as k on k.codprogramacao = a.codprogramacao and k.sequencia = 2 inner join tbtreinamentos as h on h.codtreinamento = b.codtreinamento inner join tbDadosEmpresa as i on i.codcoligada = '" & vCodcoligada & "' where a.codprogramacao = '" & Val(varGlobal) & "' order by c.nomecolaborador"
    rsRelProg.Open sqlRelInsp, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsRelProg.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsRelProg
    Screen.MousePointer = vbHourglass
    'Report.RecordSelectionFormula = "{Command.codprogramacao}= " & Val(varGlobal)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsRelProg.Close
    Set rsRelProg = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRListaPresenca.Hide
    Unload Me
    Set FCRListaPresenca = Nothing
End Sub
