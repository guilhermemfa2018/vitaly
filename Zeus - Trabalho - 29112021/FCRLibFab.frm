VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRLibFab 
   Caption         =   "Liberação de Fabricação"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRLibFab.frx":0000
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
Attribute VB_Name = "FCRLibFab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRProgramacao

Private Sub Form_Load()
On Error GoTo Err
    Dim report1 As New CRLibFab
    Dim rsLibFab As New ADODB.Recordset
    Dim sqlLibFab As String
    
    Dim crystalProgramacao As New CRAXDRT.Application
    Dim ReportProgramacao As CRAXDRT.Report
    
    If apontaLV = 19 Then vCodRel = varGlobal
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    report1.Sections("PageHeaderSection1").ReportObjects("Text1").SetText vQualquerDado(0, 30)

    If vQualquerDado(0, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text100").SetText vQualquerDado(0, 1)
    If vQualquerDado(0, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text101").SetText vQualquerDado(0, 2)
    If vQualquerDado(0, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text102").SetText vQualquerDado(0, 3)
    If vQualquerDado(0, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text103").SetText vQualquerDado(0, 4)
    If vQualquerDado(0, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text104").SetText vQualquerDado(0, 5)
    If vQualquerDado(0, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text105").SetText vQualquerDado(0, 6)
    If vQualquerDado(0, 7) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text106").SetText vQualquerDado(0, 7)
    If vQualquerDado(0, 8) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text107").SetText vQualquerDado(0, 8)
    If vQualquerDado(0, 9) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text108").SetText vQualquerDado(0, 9)
    If vQualquerDado(0, 10) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text109").SetText vQualquerDado(0, 10)
    If vQualquerDado(0, 11) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text110").SetText vQualquerDado(0, 11)
    If vQualquerDado(0, 12) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text111").SetText vQualquerDado(0, 12)
    If vQualquerDado(0, 13) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text112").SetText vQualquerDado(0, 13)
    If vQualquerDado(0, 14) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text113").SetText vQualquerDado(0, 14)
    If vQualquerDado(0, 15) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text114").SetText vQualquerDado(0, 15)
    If vQualquerDado(0, 16) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text115").SetText vQualquerDado(0, 16)
    If vQualquerDado(0, 17) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text116").SetText vQualquerDado(0, 17)
    If vQualquerDado(0, 18) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text117").SetText vQualquerDado(0, 18)
    If vQualquerDado(0, 19) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text118").SetText vQualquerDado(0, 19)
    If vQualquerDado(0, 20) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text119").SetText vQualquerDado(0, 20)
    If vQualquerDado(0, 21) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text120").SetText vQualquerDado(0, 21)
    If vQualquerDado(0, 22) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text121").SetText vQualquerDado(0, 22)
    If vQualquerDado(0, 23) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text122").SetText vQualquerDado(0, 22)
    
    rsLibFab.CursorLocation = adUseClient
    sqlLibFab = "select a.fce,d.nome,e.projeto,e.descricao,e.oc,a.datarel,a.codrel,a.observacao,a.norma,b.descposicao as item,b.desenho,b.revisao,b.qtdlib,b.posicao,b.pesolib," & _
    "f.dimensoes,b.inspsrels,a.emitidopor " & _
    "from tbRelInspExp as a inner join tbRelInspExpItens as b on a.codrel = b.codrel inner join tbFo as c on a.fce = c.fce " & _
    "inner join tbclifor as d on c.codclifor = d.codclifor inner join tbProjetos as e on a.codprojeto = e.codprojeto inner join tbItemLM as f on a.fce = f.fce and b.codlm = f.codlm and b.codseq = f.codseq where a.codrel = '" & vCodRel & "'"
    rsLibFab.Open sqlLibFab, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsLibFab.ActiveConnection = Nothing
    Set ReportProgramacao = report1
    
    ReportProgramacao.DiscardSavedData
    ReportProgramacao.Database.SetDataSource rsLibFab
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsLibFab.Close
    Set rsLibFab = Nothing
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
On Error GoTo Err
    FCRLibFab.Hide
    Set FCRLibFab = Nothing
    Unload Me
Err:
    Unload Me
End Sub


