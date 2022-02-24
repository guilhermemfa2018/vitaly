VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRProgramacao 
   Caption         =   "Programa��o"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRProgramacao.frx":0000
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
Attribute VB_Name = "FCRProgramacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRProgramacao

Private Sub Form_Load()
On Error GoTo Err
    Dim report1 As New CRProgramacao
    Dim rsProgramacao As New ADODB.Recordset
    Dim sqlProgramacao As String
    
    Dim crystalProgramacao As New CRAXDRT.Application
    Dim ReportProgramacao As CRAXDRT.Report
    
    'Linha abaixo - Altera texto do cabe�alho do relat�rio via c�digo
    report1.Sections("PageHeaderSection1").ReportObjects("Text12").SetText frmProgramacao.SkinLabel7
    
    rsProgramacao.CursorLocation = adUseClient
    sqlProgramacao = "select *,replicate('0', 2 - len(rtrim(semprog)) ) + rtrim(semprog)+replicate('0', 10 - len(rtrim(os)) ) + rtrim(os) + replicate('0', 2 - len(rtrim(revisao)) ) + rtrim(revisao) as OSRev from tbprintprogramacao order by semprog,os,revisao,idcc"
    rsProgramacao.Open sqlProgramacao, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsProgramacao.ActiveConnection = Nothing
    Set ReportProgramacao = report1
    
    ReportProgramacao.DiscardSavedData
    ReportProgramacao.Database.SetDataSource rsProgramacao
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsProgramacao.Close
    Set rsProgramacao = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRProgramacao.Hide
    Set FCRProgramacao = Nothing
    Unload Me
End Sub

