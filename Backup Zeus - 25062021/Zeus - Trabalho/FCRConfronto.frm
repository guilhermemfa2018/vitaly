VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRConfronto 
   Caption         =   "Ponto X Apropriação"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "FCRConfronto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9555
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
Attribute VB_Name = "FCRConfronto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New CRConfronto

Private Sub Form_Load()
    Dim report1 As New CRConfronto
    Dim rsOS As New ADODB.Recordset
    Dim SqlOS As String
    
    Dim crystalOS As New CRAXDRT.Application
    Dim ReportContronto As CRAXDRT.Report
    
    
    rsOS.CursorLocation = adUseClient
    
    SqlOS = "SET LANGUAGE 'Português'"
    rsOS.Open SqlOS, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    'A STRING ABAIXO REALIZA O CONFRONTO DA APROPRIAÇÃO COM O PONTO
    SqlOS = "SELECT CHAPA,NOME,CODSECAO,FUNCAO,CONVERT (VARCHAR, DATA, 103) as DATA_BATIDA,DATEPART(YEAR,DATA)AS ANO,DATEPART(MONTH,DATA)AS MESNUM,RIGHT('0'+ CONVERT(VARCHAR,cast(DATEPART(MONTH,DATA) as varchar)),2) + ' - ' + DATENAME( MONTH, DATA) as MES,DATEPART(DAY,DATA)AS DIA,CONVERT (VARCHAR, horaent, 108) as ENT_OS,CONVERT (VARCHAR, HORASAI, 108) as SAI_OS," & _
            "IDPARADA,nmparada,NOMECUSTO,REPLICATE('0', 2 - LEN(CAST((SUM(ENT1) /60) AS VARCHAR))) + CAST((SUM(ENT1) /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((SUM(ENT1) %60) AS VARCHAR))) + CAST((SUM(ENT1) %60) AS VARCHAR) AS ENT1," & _
            "REPLICATE('0', 2 - LEN(CAST((SUM(SAI1) /60) AS VARCHAR))) + CAST((SUM(SAI1) /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((SUM(SAI1) %60) AS VARCHAR))) + CAST((SUM(SAI1) %60) AS VARCHAR) AS SAI1," & _
            "REPLICATE('0', 2 - LEN(CAST((SUM(ENT2) /60) AS VARCHAR))) + CAST((SUM(ENT2) /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((SUM(ENT2) %60) AS VARCHAR))) + CAST((SUM(ENT2) %60) AS VARCHAR) AS ENT2," & _
            "REPLICATE('0', 2 - LEN(CAST((SUM(SAI2) /60) AS VARCHAR))) + CAST((SUM(SAI2) /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((SUM(SAI2) %60) AS VARCHAR))) + CAST((SUM(SAI2) %60) AS VARCHAR) AS SAI2," & _
            "REPLICATE('0', 2 - LEN(CAST((SUM(ENT3) /60) AS VARCHAR))) + CAST((SUM(ENT3) /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((SUM(ENT3) %60) AS VARCHAR))) + CAST((SUM(ENT3) %60) AS VARCHAR) AS ENT3," & _
            "REPLICATE('0', 2 - LEN(CAST((SUM(SAI3) /60) AS VARCHAR))) + CAST((SUM(SAI3) /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((SUM(SAI3) %60) AS VARCHAR))) + CAST((SUM(SAI3) %60) AS VARCHAR) AS SAI3 " & _
            "FROM (SELECT CHAPA,NOME,CODSECAO,FUNCAO,DATA,horasai,horaent,idparada,nmparada,NOMECUSTO,CASE WHEN SEQUENCIA = 1 THEN BATIDA ELSE 0 END ENT1,CASE WHEN SEQUENCIA = 2 THEN BATIDA ELSE 0 END SAI1,CASE WHEN SEQUENCIA = 3 THEN BATIDA ELSE 0 END ENT2,CASE WHEN SEQUENCIA = 4 THEN BATIDA ELSE 0 END SAI2," & _
            "CASE WHEN SEQUENCIA = 5 THEN BATIDA ELSE 0 END ENT3,CASE WHEN SEQUENCIA = 6 THEN BATIDA ELSE 0 END SAI3 FROM (SELECT ROW_NUMBER() OVER(PARTITION BY " & vBancoTotvs & ".dbo.ABATFUN.CHAPA," & vBancoTotvs & ".dbo.ABATFUN.DATA ORDER BY " & vBancoTotvs & ".dbo.ABATFUN.CHAPA, " & vBancoTotvs & ".dbo.ABATFUN.DATA, " & vBancoTotvs & ".dbo.ABATFUN.BATIDA) " & _
            "SEQUENCIA, " & vBancoTotvs & ".dbo.ABATFUN.CHAPA, " & vBancoTotvs & ".dbo.PFUNC.NOME, " & vBancoTotvs & ".dbo.PFUNC.CODSECAO," & _
            " " & vBancoTotvs & ".dbo.PFUNCAO.NOME FUNCAO, " & vBancoTotvs & ".dbo.ABATFUN.DATA,tbOsMov.horaent,tbOsMov.horasai,tbOsMov.idparada,tbParadas.nmparada, " & vBancoTotvs & ".dbo.ABATFUN.BATIDA, " & vBancoTotvs & ".dbo.GCCUSTO.NOME NOMECUSTO FROM " & vBancoTotvs & ".dbo.ABATFUN INNER JOIN " & vBancoTotvs & ".dbo.PFUNC ON " & vBancoTotvs & ".dbo.ABATFUN.CHAPA = " & vBancoTotvs & ".dbo.PFUNC.CHAPA INNER JOIN " & _
            "" & vBancoTotvs & ".dbo.PFUNCAO ON " & vBancoTotvs & ".dbo.PFUNC.CODFUNCAO =  " & vBancoTotvs & ".dbo.PFUNCAO.CODIGO " & _
            "INNER JOIN  " & vBancoTotvs & ".dbo.PFRATEIOFIXO ON  " & vBancoTotvs & ".dbo.PFUNC.CHAPA =  " & vBancoTotvs & ".dbo.PFRATEIOFIXO.CHAPA INNER JOIN  " & vBancoTotvs & ".dbo.GCCUSTO ON  " & vBancoTotvs & ".dbo.PFRATEIOFIXO.CODCCUSTO =  " & vBancoTotvs & ".dbo.GCCUSTO.CODCCUSTO left JOIN tbOsMov ON " & _
            " " & vBancoTotvs & ".dbo.ABATFUN.CHAPA = tbOsMov.chapa COLLATE SQL_Latin1_General_CP1_CI_AS and tbOsMov.dataent =  " & vBancoTotvs & ".dbo.ABATFUN.DATA " & _
            "LEFT JOIN tbParadas ON tbOsMov.idparada = tbParadas.codigo WHERE  " & vBancoTotvs & ".dbo.PFUNC.CODSITUACAO <> 'D') TBPONTOLINHA) TBPONTOCOLUNA GROUP BY CHAPA,NOME,CODSECAO,FUNCAO,DATA,horaent,horasai,IDPARADA,nmparada,NOMECUSTO order by ANO,MESNUM,NOME,DIA"
    
    rsOS.Open SqlOS, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsOS.ActiveConnection = Nothing
    Set ReportContronto = report1
    ReportContronto.DiscardSavedData
    ReportContronto.Database.SetDataSource rsOS
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
'    rsOS.Close
    Set rsOS = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRConfronto.Hide
    Unload Me
    Set FCRConfronto = Nothing
End Sub


