VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRListaColCargos 
   Caption         =   "Lista de colaboradores por cargo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRListaColCargos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   20280
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
Attribute VB_Name = "FCRListaColCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRListaColCargos
    
    Dim CRXApplication As New CRAXDDRT.Application
    Set CRXApplication = CreateObject("CrystalRuntime.Application.11")
    Dim CRXReport As New CRAXDDRT.Report
    Dim CRXDatabase As CRAXDDRT.Database
    Set CRXReport = report1
    Set CRXDatabase = CRXReport.Database
    
    
    Dim rsListaCargos As New ADODB.Recordset
    Dim SqlListaCargos As String
    
    rsListaCargos.CursorLocation = adUseClient
    SqlListaCargos = "select a.codcolaborador as REGISTRO,a.nomecolaborador,b.data,d.nomecargo,e.logo from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = 1 and  a.tipo = 'colaborador' and a.cpf = b.cpf and " & _
                     "b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo,tbDadosEmpresa as e where a.ativo = 'S' order by d.nomecargo"
    rsListaCargos.Open SqlListaCargos, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsListaCargos.ActiveConnection = Nothing
    CRXReport.DiscardSavedData
    CRXReport.Database.SetDataSource rsListaCargos
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = CRXReport
    CRViewer1.ViewReport
    CRViewer1.Zoom (100)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRListaColCargos.Hide
    Unload Me
    Set FCRListaColCargos = Nothing
End Sub

