VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRExpedicao 
   Caption         =   "Expedição"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRExpedicao.frx":0000
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
Attribute VB_Name = "FCRExpedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Err
    Dim report1 As New CRExpedicao
    Dim rsExpedicao As New ADODB.Recordset
    Dim sqlExpedicao As String
    
    Dim crystalProgramacao As New CRAXDRT.Application
    Dim ReportProgramacao As CRAXDRT.Report
    
    If apontaLV = 18 Then vCodRel = varGlobal
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    
    rsExpedicao.CursorLocation = adUseClient
    
    If vQualquerDado(20, 1) <> "-" Then
        sqlExpedicao = "select a.fce,c.pedido,d.nome,d.endereco,d.cep,d.bairro,d.cidade,d.uf,d.telefone,g.cnpj,g.inscest,e.projeto,e.descricao,e.oc,a.datarel,a.codrel,a.observacao,a.norma,b.descposicao as item," & _
                       "b.desenho,b.revisao,b.qtdlib,b.posicao,b.pesolib,a.pesobalanca,f.dimensoes,b.inspsrels,h.nome as nometransp,h.endereco as endtransp,h.cep as ceptransp,h.bairro as bairrotransp,h.cidade as cidadetransp," & _
                       "h.uf as uftransp,h.cnpj as cnpjtransp,h.inscest as ietransp,h.placaveiculo,h.uf1 as ufveiculo,h.placacarreta,h.uf2 as ufcarreta,h.motorista,a.emitidopor,b.un " & _
                       "from tbRelInspExp as a inner join tbRelInspExpItens as b on a.codrel = b.codrel inner join tbFo as c on a.fce = c.fce inner join tbclifor as d on c.codclifor = d.codclifor " & _
                       "left join tbJuridica as g on d.codclifor = g.codclifor inner join tbProjetos as e on a.codprojeto = e.codprojeto left join tbItemLM as f on a.fce = f.fce and b.codlm = f.codlm and b.codseq = f.codseq " & _
                       "left join tbRelExpTransp as h on a.codrel = h.codrel where a.codrel = '" & vCodRel & "'"
        rsExpedicao.Open sqlExpedicao, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        sqlExpedicao = "select a.fce,'-' as pedido,a.nome,a.endereco,a.cep,a.bairro,a.cidade,a.uf,a.telefone,a.cnpj,a.ie as inscest,e.projeto,e.descricao,e.oc,a.datarel,a.codrel,a.observacao,a.norma,b.descposicao as item, " & _
                    "b.desenho,b.revisao,b.qtdlib,b.posicao,b.pesolib,a.pesobalanca,f.dimensoes,b.inspsrels,h.nome as nometransp,h.endereco as endtransp,h.cep as ceptransp,h.bairro as bairrotransp,h.cidade as cidadetransp, " & _
                    "h.uf as uftransp,h.cnpj as cnpjtransp,h.inscest as ietransp,h.placaveiculo,h.uf1 as ufveiculo,h.placacarreta,h.uf2 as ufcarreta,h.motorista,a.emitidopor,b.un " & _
                    "from tbRelInspExp as a inner join tbRelInspExpItens as b on a.codrel = b.codrel left join tbProjetos as e on a.codprojeto = e.codprojeto left join tbItemLM as f on a.fce = f.fce and b.codlm = f.codlm and b.codseq = f.codseq " & _
                    "left join tbRelExpTransp as h on a.codrel = h.codrel where a.codrel = '" & vCodRel & "'"
        rsExpedicao.Open sqlExpedicao, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    Set rsExpedicao.ActiveConnection = Nothing
    Set ReportProgramacao = report1
    
    ReportProgramacao.DiscardSavedData
    ReportProgramacao.Database.SetDataSource rsExpedicao
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsExpedicao.Close
    Set rsExpedicao = Nothing
Err:
    Resume Next
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
    FCRExpedicao.Hide
    Set FCRExpedicao = Nothing
    Unload Me
Err:
    Unload Me
End Sub


