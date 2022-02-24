VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCREmprestimo 
   Caption         =   "Controle de Empréstimos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FCREmprestimo.frx":0000
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
Attribute VB_Name = "FCREmprestimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim report1 As New CREmprestimo
    Dim rsEmprestimo As New ADODB.Recordset
    Dim sqlEmprestimo As String
    
    Dim crystalEmprestimo As New CRAXDRT.Application
    Dim ReportEmprestimo As CRAXDRT.Report
    
    rsEmprestimo.CursorLocation = adUseClient
    
    
    sqlEmprestimo = "select a.chapa,a.nome,a.nomefuncao,a.nomesecao,a.codusuariorm,CONVERT (VARCHAR, b.dataemprestimo, 103) as dataemprestimo,CONVERT (VARCHAR, b.horaemprestimo, 108) as horaemprestimo,b.idmov,b.qtdemprestado-b.qtddevolvida as qtdemprestado,b.codigoprd,b.descricao,CONVERT (VARCHAR, GETDATE(), 103)  AS dataAtual,CONVERT(numeric,GETDATE()-b.dataemprestimo)-1 as dias,SUBSTRING(a.nomequememprestou,10,50),e.logo,c.codloc +' '+c.nome as localEstoque,B.numeromov,B.serie " & _
                     "from tbEmprestimo as a inner join tbEmprestimoItens as b on a.idmov = b.idmov and a.status <> 'D' and b.status <> 'D' inner join " & vBancoSAP & ".dbo.tloc as c on b.localestoque = c.CODLOC COLLATE SQL_Latin1_General_CP1_CI_AS and c.CODCOLIGADA = " & vCodColigadaRM & " and c.CODLOC = " & vLocalEstoque & " and CODFILIAL = 1,tbDadosEmpresa as e, tbParametros as f where a.chapa = '" & Mid$(varGlobal, 1, 6) & "' order by b.dataemprestimo desc,b.horaemprestimo desc"
    
    rsEmprestimo.Open sqlEmprestimo, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsEmprestimo.ActiveConnection = Nothing
    Set ReportEmprestimo = report1
    ReportEmprestimo.DiscardSavedData
    ReportEmprestimo.Database.SetDataSource rsEmprestimo
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    
    

    rsEmprestimo.Close
    'Set rsEmprestimo = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    FCREmprestimo.Hide
    Unload Me
    'Set FCREmprestimo = Nothing
End Sub


