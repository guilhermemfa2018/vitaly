VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRHistFunc 
   Caption         =   "Histórico funcional"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRHistFunc.frx":0000
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
Attribute VB_Name = "FCRHistFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim report1 As New CRHistFunc
    Dim rsHFunc As New ADODB.Recordset
    Dim SqlHFunc As String
    
    rsHFunc.CursorLocation = adUseClient
    If FiltroGeral = "Ativos" Then
        SqlHFunc = "Select a.codcolaborador,a.nomecolaborador,a.datacadastro,a.datanascimento,a.sexo,a.cpf,a.ctpsnumero,a.ctpsserie,a.cnhnumero,a.cnhtipo,a.datarecisao,a.email,a.telefone,a.celular,a.mediageral,a.foto,c.logo,d.campo1,d.campo2,d.campo3,d.campo4,d.campo5,d.id,a.observacao from " & _
                   "tbcolaboradores as a left join tbPrintHFunc as d on a.cpf = d.campo1 inner join tbDadosEmpresa as c on c.codcoligada = '" & vCodColigada & "' where a.tipo = 'colaborador' and a.ativo = 'S' order by a.cpf"
    ElseIf FiltroGeral = "Não ativos" Or FiltroGeral = "Demitidos" Then
        SqlHFunc = "Select a.codcolaborador,a.nomecolaborador,a.datacadastro,a.datanascimento,a.sexo,a.cpf,a.ctpsnumero,a.ctpsserie,a.cnhnumero,a.cnhtipo,a.datarecisao,a.email,a.telefone,a.celular,a.mediageral,a.foto,c.logo,d.campo1,d.campo2,d.campo3,d.campo4,d.campo5,d.id,a.observacao from " & _
                   "tbcolaboradores as a left join tbPrintHFunc as d on a.cpf = d.campo1 inner join tbDadosEmpresa as c on c.codcoligada = '" & vCodColigada & "' where a.tipo = 'colaborador' and a.ativo = 'N' order by a.cpf"
    ElseIf FiltroGeral = "Todos" Then
        SqlHFunc = "Select a.codcolaborador,a.nomecolaborador,a.datacadastro,a.datanascimento,a.sexo,a.cpf,a.ctpsnumero,a.ctpsserie,a.cnhnumero,a.cnhtipo,a.datarecisao,a.email,a.telefone,a.celular,a.mediageral,a.foto,c.logo,d.campo1,d.campo2,d.campo3,d.campo4,d.campo5,d.id,a.observacao from " & _
                   "tbcolaboradores as a left join tbPrintHFunc as d on a.cpf = d.campo1 inner join tbDadosEmpresa as c on c.codcoligada = '" & vCodColigada & "' where a.tipo = 'colaborador' order by a.cpf"
    End If
    rsHFunc.Open SqlHFunc, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsHFunc.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsHFunc
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (90)
    Screen.MousePointer = vbDefault
    
    rsHFunc.Close
    Set rsHFunc = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRHistFunc.Hide
    Unload Me
    Set FCRHistFunc = Nothing
End Sub

