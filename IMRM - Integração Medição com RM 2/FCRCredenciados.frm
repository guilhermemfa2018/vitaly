VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRCredenciados 
   Caption         =   "Credenciados no período"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRCredenciados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13095
   ScaleWidth      =   23880
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
Attribute VB_Name = "FCRCredenciados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRCredenciados
    Dim rsCredenciados As New ADODB.Recordset
    Dim sqlCredenciados As String
    
    Dim crystalCredenciados As New CRAXDRT.Application
    Dim ReportCredenciados As CRAXDRT.Report
    
    vdataFilter1 = frmPrintRels.DTPicker1.Value
    vdataFilter2 = frmPrintRels.DTPicker2.Value
    
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    report1.Sections("PageHeaderSection1").ReportObjects("Text1").SetText "FORNECEDORES CREDENCIADOS NO PERÍODO: " & frmPrintRels.DTPicker1.Value & " - " & frmPrintRels.DTPicker2.Value & ""
    
    rsCredenciados.CursorLocation = adUseClient
    
    
'    sqlCredenciados = "select a.idfornecedor,b.CardName as nomefornecedor,CONVERT (VARCHAR, a.datacredenciamento, 103) as datacredenciamento,a.status, c.logo from tbFornecedores as a inner join  " & vBancoSAP & ".DBO.OCRD as b on a.idfornecedor = b.CardCode COLLATE SQL_Latin1_General_CP1_CI_AS, " & _
'                      "tbDadosEmpresa as c where a.datacredenciamento BETWEEN '" & Format(vdataFilter1, "yyyy/mm/dd") & "' and '" & Format(vdataFilter2, "yyyy/mm/dd") & "' ORDER BY A.DATACREDENCIAMENTO"
    sqlCredenciados = "SET DATEFORMAT dmy select a.idfornecedor,b.CardName as nomefornecedor,CONVERT (VARCHAR, a.datacredenciamento, 103) as datacredenciamento,a.status, c.logo from tbFornecedores as a inner join  " & vBancoSAP & ".DBO.OCRD as b on a.idfornecedor = b.CardCode COLLATE SQL_Latin1_General_CP1_CI_AS, " & _
                      "tbDadosEmpresa as c where a.datacredenciamento BETWEEN '" & Format(vdataFilter1, "dd/mm/yyyy") & "' and '" & Format(vdataFilter2, "dd/mm/yyyy") & "' ORDER BY A.DATACREDENCIAMENTO"
    
    
    
    rsCredenciados.Open sqlCredenciados, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsCredenciados.RecordCount = 0 Then
        mobjMsg.Abrir "Nenhum fornecedor credenciado no período informado", Ok, critico, "Atenção"
        rsCredenciados.Close
        Set rsCredenciados = Nothing
        Unload Me
        Exit Sub
    End If
    
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
    FCRCredenciados.Hide
    Unload Me
    Set FCRCredenciados = Nothing
End Sub

