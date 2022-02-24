VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRFerEmp 
   Caption         =   "Ferramentas Emperstadas"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRFerEmp.frx":0000
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
Attribute VB_Name = "FCRFerEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRFerEmp
    Dim rsFerEmp As New ADODB.Recordset
    Dim sqlFerEmp As String
    
    Dim crystalFerEmp As New CRAXDRT.Application
    Dim ReportFerEmp As CRAXDRT.Report
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    'report1.Sections("PageHeaderSection1").ReportObjects("Text1").SetText "FORNECEDORES CREDENCIADOS NO PERÍODO: " & frmPrintRels.DTPicker1.Value & " - " & frmPrintRels.DTPicker2.Value & ""
    
    varGlobal = frmPrintRels.Text1.Text
    
    
    rsFerEmp.CursorLocation = adUseClient
    
    
    sqlFerEmp = "SELECT A.idmov,A.codigoprd,A.descricao,A.status,A.chapa,B.nome,CONVERT (VARCHAR, A.dataemprestimo, 103) as dataemprestimo,A.qtdemprestado-A.qtddevolvida AS QUANTIDADE,A.qtdemprestado,A.qtddevolvida,B.nomefuncao,e.logo FROM tbEmprestimoItens AS A " & _
                "INNER JOIN tbEmprestimo AS B ON A.idmov = B.idmov AND A.chapa = B.chapa AND A.status <> 'D',tbDadosEmpresa as e WHERE A.descricao like '%" & varGlobal & "%' and a.localestoque = " & Val(vLocalEstoque) & " order by a.descricao"
    
    rsFerEmp.Open sqlFerEmp, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsFerEmp.RecordCount = 0 Then
        mobjMsg.Abrir "Nenhuma ferramenta encontrada", Ok, critico, "Atenção"
        rsFerEmp.Close
'        FCRFerEmp.Hide
        Set rsFerEmp = Nothing
        Unload Me
        Exit Sub
    End If
    
    Set rsFerEmp.ActiveConnection = Nothing
    Set ReportFerEmp = report1
    
    ReportFerEmp.DiscardSavedData
    ReportFerEmp.Database.SetDataSource rsFerEmp
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsFerEmp.Close
    Set rsFerEmp = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    FCRFerEmp.Hide
    Unload Me
    Set FCRFerEmp = Nothing
End Sub



