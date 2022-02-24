VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRAvFornecGer 
   Caption         =   "Avaliação de Fornecedor Geral"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRAvFornecGer.frx":0000
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
Attribute VB_Name = "FCRAvFornecGer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRAvFornecGer
    Dim rsAvFornecGer As New ADODB.Recordset
    Dim sqlAvFornecGer As String
    
    Dim crystalAvFornecGer As New CRAXDRT.Application
    Dim ReportAvFornecGer As CRAXDRT.Report
    
    rsAvFornecGer.CursorLocation = adUseClient
    
    vDataFilter1 = frmPrintRels.DTPicker1.Value
    vDataFilter2 = frmPrintRels.DTPicker2.Value
    
    
'    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    report1.Sections("PageHeaderSection1").ReportObjects("Text17").SetText frmPrintRels.DTPicker1.Value
    report1.Sections("PageHeaderSection1").ReportObjects("Text11").SetText frmPrintRels.DTPicker2.Value
    
    
    report1.Sections("ReportFooterSection1").ReportObjects("Text14").SetText vQualquerDado(0, 1)
    report1.Sections("ReportFooterSection1").ReportObjects("Text15").SetText "de " & Format(vQualquerDado(0, 2), "#,##0.00;(#,##0.00)") & " à " & Format(vQualquerDado(0, 3), "#,##0.00;(#,##0.00)")
    
    report1.Sections("ReportFooterSection1").ReportObjects("Text18").SetText vQualquerDado(1, 1)
    report1.Sections("ReportFooterSection1").ReportObjects("Text19").SetText "de " & Format(vQualquerDado(1, 2), "#,##0.00;(#,##0.00)") & " à " & Format(vQualquerDado(1, 3), "#,##0.00;(#,##0.00)")
    
    report1.Sections("ReportFooterSection1").ReportObjects("Text20").SetText vQualquerDado(2, 1)
    report1.Sections("ReportFooterSection1").ReportObjects("Text21").SetText "de " & Format(vQualquerDado(2, 2), "#,##0.00;(#,##0.00)") & " à " & Format(vQualquerDado(2, 3), "#,##0.00;(#,##0.00)")
    
    sqlAvFornecGer = "select * from TempNotaFornec, tbDadosEmpresa as e"
    
    rsAvFornecGer.Open sqlAvFornecGer, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    Set rsAvFornecGer.ActiveConnection = Nothing
    Set ReportAvFornecGer = report1
    
    ReportAvFornecGer.DiscardSavedData
    ReportAvFornecGer.Database.SetDataSource rsAvFornecGer
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsAvFornecGer.Close
    Set rsAvFornecGer = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    FCRAvFornecGer.Hide
    Unload Me
    Set FCRAvFornecGer = Nothing
End Sub

