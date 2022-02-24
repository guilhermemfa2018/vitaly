VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRValidadeCertif 
   Caption         =   "Documentação de fornecedores à vencer\vencida"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRValidadeCertif.frx":0000
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
Attribute VB_Name = "FCRValidadeCertif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRValidadeCertif
    Dim rsCredenciados As New ADODB.Recordset
    Dim sqlCredenciados As String
    
    Dim crystalCredenciados As New CRAXDRT.Application
    Dim ReportCredenciados As CRAXDRT.Report
    
    rsCredenciados.CursorLocation = adUseClient
    
    sqlCredenciados = "select a.idfornecedor,a.nomefornecedor,a.nomecriterioavfornec,CONVERT (VARCHAR, a.datavalidade, 103) as validade_documento,b.logo from tbNotaAvFornec as a, tbDadosEmpresa as b where a.datavalidade is not null and GETDATE()+30 >= a.datavalidade"
    
    rsCredenciados.Open sqlCredenciados, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsCredenciados.ActiveConnection = Nothing
    Set ReportCredenciados = report1
    
    ReportCredenciados.DiscardSavedData
    ReportCredenciados.Database.SetDataSource rsCredenciados
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsCredenciados.Close
    Set rsCredenciados = Nothing

End Sub


Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    FCRValidadeCertif.Hide
    Unload Me
    Set FCRValidadeCertif = Nothing
End Sub


