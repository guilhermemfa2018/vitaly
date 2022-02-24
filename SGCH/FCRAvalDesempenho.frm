VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRAvalDesempenho 
   Caption         =   "Avaliação de desempenho"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRAvalDesempenho.frx":0000
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
Attribute VB_Name = "FCRAvalDesempenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
    Dim Report As New CRAvalDesempenho

    'A rotina abaixo ainda esta em avaliação ----------------
    Dim rsRelProg As New ADODB.Recordset
    Dim sqlRelProg As String
    Dim rsRelCorp As New ADODB.Recordset
    Dim sqlRelCorp As String
    '------------------------
    Dim crystalCAB As New CRAXDRT.Application
    Dim ReportCAB As CRAXDRT.Report
    Dim crystalCORP As New CRAXDRT.Application
    Dim ReportCORP As CRAXDRT.Report
    '------------------------

    'Teste1 -----------------
    rsRelProg.CursorLocation = adUseClient
    sqlRelProg = "Select a.codprogramacao,a.cpf,b.nomecolaborador,a.situacao,c.nometreinamento,c.objetivo,c.conteudo,c.cargahoraria,d.avaldata,d.dae,d.metodo,d.metodooutro,e.codcolaborador,e.nomeinstrutor as nomeinstrutor1,x.nomeinstrutor as nomeinstrutor2,h.logo,b.codcolaborador,g.nomecargo,a.nota,i.mediaaprovacao,i.aprovadorest,d.avalnome,d.observacao,a.obsavaliacao,c.tipo,j.codavaliacao,k.nomeavaliacao,l.pontuacao,a.codprogramacao,d.metodoA,d.metodoT,d.metodoS,d.metodoPT from " & _
                 "tbpendentescur as a inner join tbcolaboradores as b on b.cpf=a.cpf inner join tbtreinamentos as c on c.codtreinamento = a.codtreinamento inner join tbprogramacao as d on d.codprogramacao=a.codprogramacao left join tbProgramacaoInstrutores as e on e.codprogramacao = d.codprogramacao and e.sequencia = 1 left join tbProgramacaoInstrutores as x on x.codprogramacao = d.codprogramacao and x.sequencia = 2 inner join tbmatriz as f on a.codmatriz = f.codmatriz " & _
                 "inner join tbavaliacaoprog as j on d.codmodelo = j.codmodelo left join tbavaliacaotrei as l on a.codprogramacao = l.codprogramacao and a.cpf = l.cpf and j.codavaliacao = l.codavaliacao inner join tbavaliacao as k on j.codavaliacao = k.codavaliacao inner join tbcargos as g on f.codcargo = g.codcargo inner join tbDadosEmpresa as h on h.codcoligada ='" & vCodcoligada & "',tbparametros as i " & _
                 "Where b.ativo = 'S' Order by a.cpf,j.codavaliacao"
'    sqlRelProg = "select a.codprogramacao,a.cpf,b.nomecolaborador,a.situacao,c.nometreinamento,c.objetivo,c.cargahoraria,d.avaldata,d.dae,d.metodo,d.metodooutro,e.codcolaborador,e.nomeinstrutor,h.logo,b.codcolaborador,g.nomecargo,a.nota,i.mediaaprovacao,i.aprovadorest,d.avalnome,d.observacao,a.obsavaliacao,c.tipo,j.codavaliacao,k.nomeavaliacao,a.codprogramacao from " & _
'                 "tbpendentescur as a inner join tbcolaboradores as b on b.cpf=a.cpf inner join tbtreinamentos as c on c.codtreinamento = a.codtreinamento inner join tbprogramacao as d on d.codprogramacao=a.codprogramacao " & _
'                 "left join tbProgramacaoInstrutores as e on e.codprogramacao = d.codprogramacao inner join tbmatriz as f on a.codmatriz = f.codmatriz inner join tbcargos as g on f.codcargo = g.codcargo, tbDadosEmpresa as h, tbparametros as i, tbavaliacaoprog as j inner join tbavaliacao as k on j.codavaliacao = k.codavaliacao where j.codmodelo = d.codmodelo order by b.nomecolaborador"
    rsRelProg.Open sqlRelProg, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsRelProg.ActiveConnection = Nothing
    Set ReportCAB = Report
    ReportCAB.DiscardSavedData
    ReportCAB.Database.SetDataSource rsRelProg
    Screen.MousePointer = vbHourglass

    Report.RecordSelectionFormula = "{ADParticipantes.codprogramacao}= " & Val(varGlobal)

    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault

    rsRelProg.Close
    Set rsRelProg = Nothing

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
    FCRAvalDesempenho.Hide
    Unload Me
    Set FCRAvalDesempenho = Nothing
End Sub
