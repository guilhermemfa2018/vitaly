VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form FCRNaoFechou 
   Caption         =   "Relatório de Colaboradores que não fecharam OS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRNaoFechou.frx":0000
   LinkTopic       =   "Form2"
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
Attribute VB_Name = "FCRNaoFechou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim report1 As New CRNaoFechou
    Dim rsNaoFechou As New ADODB.Recordset
    Dim sqlNaoFechou As String
    
    Dim crystalNaoFechou As New CRAXDRT.Application
    Dim ReportNaoFechou As CRAXDRT.Report
    
    rsNaoFechou.CursorLocation = adUseClient
    sqlNaoFechou = "select id,CONVERT (VARCHAR, dia, 103) as dia,registro,nome,os,fechamento,CONVERT (VARCHAR, GETDATE(), 103) as dataatual,nomecc,idoperacao from tbNaoFechaOs where CONVERT (VARCHAR, dia, 103) = CONVERT (VARCHAR, GETDATE()-1, 103)"
'    sqlNaoFechou = "select id,CONVERT (VARCHAR, dia, 103) as dia,registro,nome,os,fechamento,CONVERT (VARCHAR, GETDATE(), 103) as dataatual from tbNaoFechaOs where CONVERT (VARCHAR, dia, 103) = CONVERT (VARCHAR, GETDATE(), 103)"
    rsNaoFechou.Open sqlNaoFechou, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsNaoFechou.ActiveConnection = Nothing
    Set ReportNaoFechou = report1
    ReportNaoFechou.DiscardSavedData
    ReportNaoFechou.Database.SetDataSource rsNaoFechou
    Screen.MousePointer = vbHourglass
    'Report.RecordSelectionFormula = "{OrdemServico.os}= " & Val(varGlobal)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsNaoFechou.Close
    Set rsNaoFechou = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRNaoFechou.Hide
    Unload Me
    Set FCRNaoFechou = Nothing
End Sub




