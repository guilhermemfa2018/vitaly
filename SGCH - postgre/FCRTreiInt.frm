VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRTreiInt 
   Caption         =   "Treinamentos introdutórios realizados"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FCRTreiInt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
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
Attribute VB_Name = "FCRTreiInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRTreiInt
    Dim rsTreiInt As New ADODB.Recordset
    Dim SqlTreiInt As String
    
    rsTreiInt.CursorLocation = adUseClient
    SqlTreiInt = "select a.cpf,b.codcolaborador,b.nomecolaborador,a.codmatriz,g.nomecargo,g.codcbo,e.data,h.nomedepartamento,i.nomesetor,f.nivel,a.codtreinamento,c.nometreinamento,c.origem,c.conteudo,c.objetivo,c.introdutorio,a.codprogramacao,a.ativo,a.status,d.avaldata,f.codmatriz,d.codcolaborador,j.nomecolaborador,k.logo,b.foto from " & _
                "tbPendentesCur as a inner join tbcolaboradores as b on a.cpf = b.cpf inner join tbtreinamentos as c on a.codtreinamento = c.codtreinamento inner join tbprogramacao as d on d.codprogramacao = a.codprogramacao inner join tbcolaboradoreshist as e on a.cpf = e.cpf inner join tbmatriz as f on f.codmatriz = e.codmatriz " & _
                "inner join tbcargos as g on f.codcargo = g.codcargo inner join tbdepartamentos as h on f.coddepartamento = h.coddepartamento inner join tbsetores as i on f.codsetor = i.codsetor inner join tbcolaboradores as j on d.codcolaborador = j.codcolaborador inner join tbDadosEmpresa as k on k.codcoligada = '" & vCodColigada & "' where c.introdutorio = 'S' " & _
                "and a.status = 'Concluido' and  a.ativo = 'S' order by  a.codprogramacao"
    rsTreiInt.Open SqlTreiInt, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsTreiInt.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsTreiInt
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (75)
    Screen.MousePointer = vbDefault
    
    rsTreiInt.Close
    Set rsTreiInt = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRTreiInt.Hide
    Unload Me
    Set FCRTreiInt = Nothing
End Sub

