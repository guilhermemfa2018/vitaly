VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRListaCargos 
   Caption         =   "Lista de cargos ativos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRListaCargos.frx":0000
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
Attribute VB_Name = "FCRListaCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'    Dim report1 As New CRListaCargos
'    Dim rsListaCargos As New ADODB.Recordset
'    Dim SqlListaCargos As String
    
'    rsListaCargos.CursorLocation = adUseClient
'    SqlListaCargos = "select a.*, h.logo from tbcargos as a,tbDadosEmpresa as h where ativo = 'S'"
'    rsListaCargos.Open SqlListaCargos, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
'    Set rsListaCargos.ActiveConnection = Nothing
'    Set Report = report1
'    Report.DiscardSavedData
'    Report.Database.SetDataSource rsListaCargos
'    Screen.MousePointer = vbHourglass

'    CRViewer1.ReportSource = report1
'    CRViewer1.ViewReport
'    CRViewer1.Zoom (100)
'    Screen.MousePointer = vbDefault
    
'    rsListaCargos.Close
'    Set rsListaCargos = Nothing
'    Exit Sub
'------------------------
    Dim report1 As New CRListaCargos
    
    Dim CRXApplication As New CRAXDDRT.Application
    Set CRXApplication = CreateObject("CrystalRuntime.Application.11")
    Dim CRXReport As New CRAXDDRT.Report
    Dim CRXDatabase As CRAXDDRT.Database
    Set CRXReport = report1
    Set CRXDatabase = CRXReport.Database
    
    
    Dim rsListaCargos As New ADODB.Recordset
    Dim SqlListaCargos As String
    
    rsListaCargos.CursorLocation = adUseClient
    SqlListaCargos = "select a.*, h.logo from tbcargos as a inner join tbDadosEmpresa as h on h.codcoligada = '" & vCodColigada & "' where ativo = 'S' order by a.nomecargo"
    rsListaCargos.Open SqlListaCargos, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsListaCargos.ActiveConnection = Nothing
    CRXReport.DiscardSavedData
    CRXReport.Database.SetDataSource rsListaCargos
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = CRXReport
    CRViewer1.ViewReport
    CRViewer1.Zoom (100)
    Screen.MousePointer = vbDefault
'--------------------
'    Dim Report As New CRListaCargos

'    Screen.MousePointer = vbHourglass
'    CRViewer1.ReportSource = Report
'    CRViewer1.ViewReport
'    Screen.MousePointer = vbDefault


End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRListaCargos.Hide
    Unload Me
    Set FCRListaCargos = Nothing
End Sub
