VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRLibparaInsp 
   Caption         =   "Liberados Para Inspeção em Preto"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRLibparaInsp.frx":0000
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
Attribute VB_Name = "FCRLibparaInsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim report1 As New CRLibparaInsp
    Dim rsLibparaInsp As New ADODB.Recordset
    Dim SqlLibparaInsp As String
    
    Dim crystalOS As New CRAXDRT.Application
    Dim ReportOS As CRAXDRT.Report
    
    
    rsLibparaInsp.CursorLocation = adUseClient
'    SqlLibparaInsp = "select a.idos,a.revisaoos,a.idoperacao,a.idcc,a.nomecc,a.codigobarra,C.fce,a.status as status_lib,d.idcc,d.status as status_acab,b.desenho,e.revisao as revisao_des from tbMPItens AS A INNER JOIN tbMP as B ON A.idprogramacao = B.idprogramacao and a.idcc like '7000.710%' inner join tbMPItens AS D " & _
'                     "ON d.idprogramacao = B.idprogramacao and d.idcc = '3000.3105.SC-01' inner join tbProjetos as C on B.codprojeto = C.codprojeto inner join tbdesenhos as e on b.desenho = e.desenho where d.status = 3 and a.status<=2 order by a.idos,a.idoperacao desc"
    
    SqlLibparaInsp = "select a.idos,a.revisaoos,a.idoperacao,a.idcc,a.nomecc,a.codigobarra,C.fce,a.status as status_lib,d.idcc,d.status as status_acab,g.desenho,g.revisao as revisao_des from tbMPItens AS A INNER JOIN tbMP as B ON A.idprogramacao = B.idprogramacao and a.idcc like '7000.710%' inner join tbMPItens AS D ON d.idprogramacao = B.idprogramacao and d.idcc = '3000.3105.SC-01' " & _
                     "inner join tbProjetos as C on B.codprojeto = C.codprojeto left join tbitemlm as f on SUBSTRING(a.desenhos,1,2) = f.codlm and replace(SUBSTRING(a.desenhos,3,4),';','') = f.codseq and c.fce = f.fce left join tbDesenhos as g on f.codigodes = g.iddesenho where d.status = 3 and a.status<=2 order by a.idos,a.idoperacao desc"
        
    rsLibparaInsp.Open SqlLibparaInsp, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsLibparaInsp.ActiveConnection = Nothing
    Set ReportOS = report1
    ReportOS.DiscardSavedData
    ReportOS.Database.SetDataSource rsLibparaInsp
    Screen.MousePointer = vbHourglass
    'Report.RecordSelectionFormula = "{OrdemServico.os}= " & Val(varGlobal)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsLibparaInsp.Close
    Set rsLibparaInsp = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRLibparaInsp.Hide
    Unload Me
    Set FCRLibparaInsp = Nothing
End Sub

