VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCREncargosSalarios 
   Caption         =   "Relatório de Encargos e Salários"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCREncargosSalarios.frx":0000
   LinkTopic       =   "Form2"
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
Attribute VB_Name = "FCREncargosSalarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CREncargosSalarios
Private Sub Form_Load()
    Dim report1 As New CREncargosSalarios
    Dim rsEncargosSalarios As New ADODB.Recordset
    Dim sqlEncargosSalarios As String
    
    rsEncargosSalarios.CursorLocation = adUseClient
    
    sqlEncargosSalarios = "SELECT A.IDLAN AS REFERENCIA,A.CODTDO AS TIPO_DOC,B.DESCRICAO,D.CODREDUZIDO,D.NOME,C.VALOR,CONVERT (VARCHAR, A.DATAEMISSAO, 103) as DATA,DATEPART(YEAR,A.DATAEMISSAO)AS ANO," & _
                     "DATEPART(MONTH,A.DATAEMISSAO)AS MES,DATEPART(DAY,A.DATAEMISSAO)AS DIA,A.HISTORICO,C.CODCCUSTO from FLAN AS A INNER JOIN PLANCFINANC AS B ON A.CODTDO = B.CODTDO AND A.CODTDO IN(" & vMovs & ") and " & _
                     "A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' and A.CODCOLIGADA = '" & vColigada & "' INNER JOIN FLANRATCCU AS C ON A.IDLAN = C.IDLAN  AND C.CODCOLIGADA = '" & vColigada & "'  INNER JOIN GCCUSTO " & _
                     "AS D ON C.CODCCUSTO = D.CODCCUSTO AND A.CODCOLIGADA=D.CODCOLIGADA GROUP BY A.IDLAN,A.CODTDO,B.DESCRICAO,D.CODREDUZIDO,D.NOME,C.VALOR,CONVERT (VARCHAR, A.DATAEMISSAO, 103),DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO),DATEPART(DAY,A.DATAEMISSAO),A.HISTORICO,C.CODCCUSTO ORDER BY ANO,MES,DIA,REFERENCIA,TIPO_DOC,C.CODCCUSTO"
    
    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsEncargosSalarios.Open sqlEncargosSalarios, cnBanco, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsEncargosSalarios.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsEncargosSalarios
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (166)
    Screen.MousePointer = vbDefault
    
    rsEncargosSalarios.Close
    Set rsEncargosSalarios = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCREncargosSalarios.Hide
    Unload Me
    Set FCREncargosSalarios = Nothing
    Form1.MousePointer = 0
    Form1.Label1.Visible = False
End Sub








