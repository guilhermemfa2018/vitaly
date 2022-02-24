VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCROrdemServico 
   Caption         =   "Ordem de Servi�o"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCROrdemServico.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
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
Attribute VB_Name = "FCROrdemServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New CROrdemServico

Private Sub Form_Load()
    Dim report1 As New CROrdemServico
    Dim rsOS As New ADODB.Recordset
    Dim SqlOS As String
    
    Dim crystalOS As New CRAXDRT.Application
    Dim ReportOS As CRAXDRT.Report
    
    
    rsOS.CursorLocation = adUseClient
    SqlOS = "Declare @TempoTotal as VARCHAR(40)SET @TempoTotal = '' " & _
             "Declare @Desenhos as VARCHAR(100) SET @Desenhos = '' " & _
             "select @Desenhos = @Desenhos + d.desenho + ' / ' from tbos as a inner join tbOsItens as b on a.idos = b.idos left join tbItemLM as c on b.fce = c.fce and b.codlm = c.codlm and " & _
             "b.codseq = c.codseq left join tbdesenhos as d on c.codigodes = d.iddesenho where b.fce = '" & vFCE & "' and b.idprogramacao = '" & varGlobal & "' group by d.desenho " & _
             "" & _
             "SELECT @TempoTotal = dbo.FN_CONVMIN(sum((cast(replace(a.tempocalc,'.','') as money)/100))) " & _
             "from tbMPItens as a where a.idprogramacao = '" & varGlobal & "' " & _
             "select a.idos,f.revisaoos,b.idprogramacao,d.descricao,a.dataos,f.dataprevista,a.rastreabilidade,b.fce,g.projeto,@Desenhos as Desenhos,d.revisao,e.posicao,e.item,'Horas' = dbo.FN_CONVMIN((cast(replace(f.tempocalc,'.','') as money)/100)),h.responsavel,a.observacao,f.codigobarra,f.idcc,f.nomecc,f.observacao," & _
             "f.idoperacao,j.nome as nomecli,l.observacao2,@TempoTotal as TempoTotal from tbos as a inner join tbOsItens as b on a.idos = b.idos left join tbItemLM as c on b.fce = c.fce and b.codlm = c.codlm and b.codseq = c.codseq left join tbdesenhos as d on c.codigodes = d.iddesenho " & _
             "left join tbPosicoes as e on c.codigodes = e.codigodes and c.codigopos = e.codigopos inner join tbMPItens as f on b.idprogramacao = f.idprogramacao and a.idos = f.idos and a.revisao = f.revisaoos " & _
             "left join tbProjetos as g on d.codprojeto = g.codprojeto inner join tbMP as h on b.idprogramacao = h.idprogramacao inner join tbFo as i on b.fce = i.fce inner join tbclifor as j on i.codclifor = j.codclifor inner join tbFormula as l on b.idcc = l.codreduzido and l.idform = 1 where b.fce = '" & vFCE & "' and b.idprogramacao = '" & varGlobal & "'"
    
    rsOS.Open SqlOS, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsOS.ActiveConnection = Nothing
    Set ReportOS = report1
    ReportOS.DiscardSavedData
    ReportOS.Database.SetDataSource rsOS
    Screen.MousePointer = vbHourglass
    'Report.RecordSelectionFormula = "{OrdemServico.os}= " & Val(varGlobal)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsOS.Close
    Set rsOS = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCROrdemServico.Hide
    Unload Me
    Set FCROrdemServico = Nothing
End Sub


