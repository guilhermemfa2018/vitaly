VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRADP 
   Caption         =   "ADP - Avaliação de Desempenho Profissional"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRADP.frx":0000
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
Attribute VB_Name = "FCRADP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRADP

Private Sub Form_Load()
'Screen.MousePointer = vbHourglass
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault
'CRViewer1.Zoom (120)
    
    Dim report1 As New CRADP
    Dim rsADP As New ADODB.Recordset
    Dim sqlADP As String
    
    rsADP.CursorLocation = adUseClient
    sqlADP = "select a.id,d.cpf,a.codcolaborador,d.nomecolaborador,h.nomecargo,e.data as admissao,a.tipoADP,a.dias,a.dataavaliacao,a.datavencimento,a.datadevolucao,a.codrespADP,a.nomerespADP,a.ausenciaano,a.atrasoano,a.codrespABS," & _
             "a.nomerespABS,a.observacao,a.indicacaotipo,a.indicacaomod1,a.indicacaomod2,a.indicacaomod3,a.indicacaomod4,a.indicacaomod5,a.indicacaomod6,a.indicacaooutros,a.nota as [Media Geral],b.codavaliacao,c.nomeavaliacao," & _
             "b.nota as [Nota Item],b.dimensao,c.descricao,d.foto,f.logo,i.mediaaprovacao,i.aprovadorest,j.pontos as [Ausencia],k.pontos as [Atraso] " & _
             "from tbListaADP as a inner join tbListaADPItens as b on a.id = b.idADP AND a.codcoligada = '" & vCodColigada & "' left join tbavaliacao as c on b.codavaliacao = c.codavaliacao " & _
             "inner join tbColaboradores as d on a.codcolaborador = d.id inner join tbcolaboradoreshist as e on d.cpf=e.cpf and e.ativo = 'S' " & _
             "inner join tbmatriz as g on e.codmatriz = g.codmatriz left join tbABS as j on j.tipo = 'Ausência' and a.ausenciaano >= j.oc1 and " & _
             "a.ausenciaano < j.oc2 left join tbABS as k on k.tipo = 'Atraso' and a.atrasoano >= k.oc1 and a.atrasoano < k.oc2 inner join tbcargos as h " & _
             "on g.codcargo = h.codcargo inner join tbDadosEmpresa as f on a.codcoligada = f.codcoligada, tbparametros as i"
    rsADP.Open sqlADP, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsADP.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsADP
    Screen.MousePointer = vbHourglass
    Report.RecordSelectionFormula = "{ADP.id}= " & Val(varGlobal2)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsADP.Close
    Set rsADP = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRADP.Hide
    Unload Me
    Set FCRADP = Nothing
End Sub

