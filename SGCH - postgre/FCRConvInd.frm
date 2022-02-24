VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRConvInd 
   Caption         =   "Convocação de treinamento individual"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   Icon            =   "FCRConvInd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      Height          =   9405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _cx             =   17171
      _cy             =   16589
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
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
Attribute VB_Name = "FCRConvInd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRConvInd

Private Sub Form_Load()
    Dim report1 As New CRConvInd
    Dim rsConvocacao As New ADODB.Recordset
    Dim SqlConvocacao As String
    
    rsConvocacao.CursorLocation = adUseClient
    SqlConvocacao = "select * from ##tbColabsConvocados as a inner join tbConfConvocacao as b on a.ID=b.ID,tbDadosEmpresa as h where h.codcoligada  = '" & vCodColigada & "'"
    rsConvocacao.Open SqlConvocacao, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsConvocacao.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsConvocacao
    Screen.MousePointer = vbHourglass

    CrystalActiveXReportViewer1.ReportSource = report1
    CrystalActiveXReportViewer1.ViewReport
    CrystalActiveXReportViewer1.Zoom (100)
    Screen.MousePointer = vbDefault
    
    rsConvocacao.Close
    Set rsConvocacao = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CrystalActiveXReportViewer1.Top = 0
    CrystalActiveXReportViewer1.Left = 0
    CrystalActiveXReportViewer1.Height = ScaleHeight
    CrystalActiveXReportViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRConvInd.Hide
    Unload Me
    Set FCRConvInd = Nothing
End Sub
