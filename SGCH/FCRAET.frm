VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRAET 
   Caption         =   "Avaliação da Eficácia de Treinamento"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRAET.frx":0000
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
      EnableGroupTree =   0   'False
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
Attribute VB_Name = "FCRAET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRAET

Private Sub Form_Load()
    Dim report1 As New CRAET
    Dim rsRelProg As New ADODB.Recordset
    Dim sqlRelProg As String
    
    rsRelProg.CursorLocation = adUseClient
    sqlRelProg = "select a.codprogramacao,a.cpf,b.nomecolaborador,a.situacao,c.nometreinamento,c.objetivo,c.cargahoraria,d.avaldata,d.dae,d.metodo,d.metodooutro,e.codcolaborador,e.nomeinstrutor as nomeinstrutor1,x.nomeinstrutor as nomeinstrutor2,h.logo,b.codcolaborador,g.nomecargo,a.nota,i.mediaaprovacao,i.aprovadorest,d.avalnome,d.observacao,a.obsavaliacao,c.tipo,d.metodoA,d.metodoT,d.metodoS,d.metodoPT from " & _
                 "tbpendentescur as a inner join tbcolaboradores as b on  b.cpf=a.cpf inner join  tbtreinamentos as c on  c.codtreinamento = a.codtreinamento inner join  tbprogramacao as d on d.codprogramacao=a.codprogramacao left join tbProgramacaoInstrutores as e on e.codprogramacao = d.codprogramacao  and e.sequencia = 1 left join tbProgramacaoInstrutores as x on x.codprogramacao = d.codprogramacao and x.sequencia = 2 inner join tbmatriz as f on a.codmatriz = f.codmatriz inner join tbcargos as g on f.codcargo = g.codcargo inner join tbDadosEmpresa as h on h.codcoligada = '" & vCodcoligada & "', tbparametros as i where b.ativo = 'S' order by b.nomecolaborador"
    rsRelProg.Open sqlRelProg, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsRelProg.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsRelProg
    Screen.MousePointer = vbHourglass
    Report.RecordSelectionFormula = "{AET.codprogramacao}= " & Val(varGlobal)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsRelProg.Close
    Set rsRelProg = Nothing
'    FCRAET.Hide
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRAET.Hide
    Unload Me
    Set FCRAET = Nothing
End Sub

