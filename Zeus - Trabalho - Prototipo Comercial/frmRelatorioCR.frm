VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form frmRelatorioCR 
   Caption         =   "Relatorios"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmRelatorioCR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   9045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      _cx             =   15531
      _cy             =   15954
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
Attribute VB_Name = "frmRelatorioCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim Report1 As New CrystalReport1
    Dim rsProgramacao As New ADODB.Recordset
    Dim sqlProgramacao As String
    
    Dim crystalProgramacao As New CRAXDRT.Application
    Dim ReportProgramacao As CRAXDRT.Report
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    
    sqlProgramacao = "select a.fce,d.nome,e.projeto,e.descricao,e.oc,a.datarel,a.codrel,b.descposicao as item,b.desenho,b.revisao,b.qtdlib,b.posicao,b.pesolib from tbRelInspExp as a inner join tbRelInspExpItens as b on a.codrel = b.codrel inner join tbFo as c on a.fce = c.fce " & _
    "inner join tbclifor as d on c.codclifor = d.codclifor inner join tbProjetos as e on a.codprojeto = e.codprojeto where a.codrel = '" & Val(Mid$(varGlobal, 1, 6)) & "'"
    rsProgramacao.Open sqlProgramacao, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsProgramacao.ActiveConnection = Nothing
    Set ReportProgramacao = Report1
    
    ReportProgramacao.DiscardSavedData
    ReportProgramacao.Database.SetDataSource rsProgramacao
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsProgramacao.Close
    Set rsProgramacao = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRelatorioCR.Hide
    Unload Me
    Set frmRelatorioCR = Nothing
End Sub

