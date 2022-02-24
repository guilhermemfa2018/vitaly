VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form FCRTreinCargo 
   Caption         =   "Cargos por treinamento"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRTreiCargo.frx":0000
   KeyPreview      =   -1  'True
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
Attribute VB_Name = "FCRTreinCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRTreinCargo

Private Sub Form_Load()

    Dim report1 As New CRTreinCargo
    Dim rsTreinCargo As New ADODB.Recordset
    Dim SqlTreinCargo As String
    
    
    rsTreinCargo.CursorLocation = adUseClient
    SqlTreinCargo = "select a.codtreinamento, a.nometreinamento, b.codsetor, c.codcargo, d.nomecargo, h.logo,substring(a.objetivo,1,300),substring(a.conteudo,1,300) from tbtreinamentos as a inner join tbTreinamentosInt as b on a.codtreinamento = b.codtreinamento " & _
                    "left join tbmatriz as c on b.codsetor = c.codsetor left join tbcargos as d on c.codcargo = d.codcargo inner join tbDadosEmpresa as h on h.codcoligada = '" & vCodcoligada & "' group by a.codtreinamento, a.nometreinamento, b.codsetor,c.codcargo," & _
                    "d.nomecargo,h.logo,substring(a.objetivo,1,300),substring(a.conteudo,1,300) order by a.codtreinamento,c.codcargo desc"
    rsTreinCargo.Open SqlTreinCargo, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsTreinCargo.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsTreinCargo
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (102)
    Screen.MousePointer = vbDefault
    
    rsTreinCargo.Close
    Set rsTreinCargo = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
