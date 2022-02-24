VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRConvocacao 
   Caption         =   "Convoca��o de treinamento"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRConvocacao.frx":0000
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
Attribute VB_Name = "FCRConvocacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRConvocacao

Private Sub Form_Load()
    Dim report1 As New CRConvocacao
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

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (80)
    Screen.MousePointer = vbDefault
    
    rsConvocacao.Close
    Set rsConvocacao = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRConvocacao.Hide
    Unload Me
    Set FCRConvocacao = Nothing
End Sub

