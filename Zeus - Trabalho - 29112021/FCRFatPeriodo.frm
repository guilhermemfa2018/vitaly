VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRFatPeriodo 
   Caption         =   "Faturamento por Período"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRFatPeriodo.frx":0000
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
Attribute VB_Name = "FCRFatPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err
    Dim report1 As New CRFatPeriodo
    Dim rsFatPeriodo As New ADODB.Recordset
    Dim sqlFatPeriodo As String
    
    Dim crystalProgramacao As New CRAXDRT.Application
    Dim ReportProgramacao As CRAXDRT.Report
    
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    report1.Sections("PageHeaderSection1").ReportObjects("Text8").SetText "RELATÓRIO DE FATURAMENTO NO PERÍODO: " & frmPrintRels.DTPicker1.Value & " - " & frmPrintRels.DTPicker2.Value & ""
    
    rsFatPeriodo.CursorLocation = adUseClient
    
'    sqlFatPeriodo = "SELECT B.CODTB3FAT,CONVERT (VARCHAR, B.DATAEMISSAO, 103) as DATAEMISSAO,B.IDMOV,B.NUMEROMOV,MAX(B.PESOBRUTO) AS PESO,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER,CASE WHEN B.STATUS = 'P' THEN 'Parcialmente Quitado' " & _
'                   "WHEN B.STATUS = 'C' THEN 'Cancelado' WHEN B.STATUS = 'A' THEN 'Pendente/Faturar' WHEN B.STATUS = 'Q' THEN 'Quitado' WHEN B.STATUS = 'F' THEN 'Receber/A pagar' ELSE 'NÃO IDENTIFICADO' END AS STATUS,CASE WHEN B.CODTMV = '2.2.01' OR B.CODTMV = '2.2.05' THEN 'FATURAMENTO' " & _
'                   "WHEN B.CODTMV = '1.2.15' OR B.CODTMV = '1.2.17' THEN 'DEVOLUÇÃO' ELSE 'ADIANTAMENTO' END CODTMV FROM CORPORERM.dbo.TMOV AS B INNER JOIN CORPORERM.dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05','2.2.25','1.2.15','1.2.17') AND B.STATUS <> 'C' " & _
'                   "WHERE B.CODTB3FAT is not null and B.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' GROUP BY B.CODTB3FAT,B.IDMOV,B.DATAEMISSAO,B.NUMEROMOV,B.STATUS,B.CODTMV order by b.CODTB3FAT"
'    sqlFatPeriodo = "SELECT T1.CODTB3FAT,T1.DATAEMISSAO,T1.IDMOV,T1.NUMEROMOV,T1.PESO,T1.VALOR_ORIGINAL,T1.VALOR_BAIXADO,T1.VALOR_RECEBER,T1.STATUS,T1.CODTMV,T2.PESO_TOTAL AS PESO_TOTAL_COMERCIAL,T2.VALOR_TOTAL AS VALOR_TOTAL_COMERCIAL,T3.PESO AS PESO_TOTAL_GLOBAL,T3.VALOR_ORIGINAL AS VALOR_TOTAL_GLOBAL,T2.nome AS NOME_CLI,T2.dataentrega " & _
'                    "FROM (SELECT B.CODTB3FAT,CONVERT (VARCHAR, B.DATAEMISSAO, 103) as DATAEMISSAO,B.IDMOV,B.NUMEROMOV,MAX(B.PESOBRUTO) AS PESO,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER,CASE WHEN B.STATUS = 'P' THEN 'Parcialmente Quitado' " & _
'                    "WHEN B.STATUS = 'C' THEN 'Cancelado' WHEN B.STATUS = 'A' THEN 'Pendente/Faturar' WHEN B.STATUS = 'Q' THEN 'Quitado' WHEN B.STATUS = 'F' THEN 'Receber/A pagar' ELSE 'NÃO IDENTIFICADO' END AS STATUS,CASE WHEN B.CODTMV = '2.2.01' OR B.CODTMV = '2.2.05' THEN 'FATURAMENTO' WHEN B.CODTMV = '1.2.15' OR B.CODTMV = '1.2.17' THEN 'DEVOLUÇÃO' " & _
'                    "ELSE 'ADIANTAMENTO' END CODTMV FROM CORPORERM.dbo.TMOV AS B INNER JOIN CORPORERM.dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05','2.2.25','1.2.15','1.2.17') AND B.STATUS <> 'C' WHERE B.CODTB3FAT is not null and B.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' GROUP BY B.CODTB3FAT,B.IDMOV,B.DATAEMISSAO,B.NUMEROMOV,B.STATUS,B.CODTMV) T1 " & _
'                    "LEFT JOIN (SELECT A.FCE,d.nome,CONVERT (VARCHAR, b.dataentrega, 103) as dataentrega,SUM(A.PESO) AS PESO_TOTAL,SUM(A.TOTAL) AS VALOR_TOTAL FROM TBPEDIDOS AS A inner join tbFCE as b on a.fce = b.fce inner join tbFo as c on b.fce = c.fce inner join tbclifor as d on c.codclifor = d.codclifor GROUP BY a.fce,d.nome,b.dataentrega) T2 ON T1.CODTB3FAT = T2.FCE LEFT JOIN " & _
'                    "(SELECT B.CODTB3FAT AS FCE,SUM(B.PESOBRUTO) AS PESO,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL FROM CORPORERM.dbo.TMOV AS B INNER JOIN CORPORERM.dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05','2.2.25','1.2.15','1.2.17') AND B.STATUS <> 'C' WHERE B.CODTB3FAT is not null GROUP BY B.CODTB3FAT) T3 ON T1.CODTB3FAT = T3.FCE"
    
    sqlFatPeriodo = "SELECT T1.CODTB3FAT,T1.DATAEMISSAO,T1.IDMOV,T1.NUMEROMOV,T1.NUMERODOCUMENTO,T1.DATAVENCIMENTO,T1.HISTORICO,T1.PESO,T1.VALOR_ORIGINAL,T1.VALOR_BAIXADO,T1.VALOR_RECEBER,T1.STATUS,T1.CODTMV,T2.PESO_TOTAL AS PESO_TOTAL_COMERCIAL,T2.VALOR_TOTAL AS VALOR_TOTAL_COMERCIAL,T3.PESO AS PESO_TOTAL_GLOBAL,T3.VALOR_ORIGINAL AS VALOR_TOTAL_GLOBAL,T2.nome AS NOME_CLI,T2.dataentrega,T2.databook,T2.adiantamento,T1.DATABAIXA,T3.VALOR_RECEBER AS VALOR_RECEBER_GLOBAL " & _
                    "FROM (SELECT B.CODTB3FAT,B.DATAEMISSAO,B.IDMOV,B.NUMEROMOV,(B.PESOBRUTO) AS PESO,(C.VALORORIGINAL) AS VALOR_ORIGINAL,(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER,CASE WHEN B.STATUS = 'P' THEN 'Parcialmente Quitado' WHEN B.STATUS = 'C' THEN 'Cancelado' WHEN B.STATUS = 'A' THEN 'Pendente/Faturar' WHEN B.STATUS = 'Q' THEN 'Quitado' " & _
                    "WHEN B.STATUS = 'F' THEN 'Receber/A pagar' ELSE 'NÃO IDENTIFICADO' END AS STATUS,CASE WHEN B.CODTMV = '2.2.01' OR B.CODTMV = '2.2.05' THEN 'FATURAMENTO' WHEN B.CODTMV = '1.2.15' OR B.CODTMV = '1.2.17' THEN 'DEVOLUÇÃO' ELSE 'ADIANTAMENTO' END CODTMV,C.NUMERODOCUMENTO,C.HISTORICO,C.DATAVENCIMENTO,C.DATABAIXA FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = " & _
                    "C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05','2.2.25','1.2.15','1.2.17') AND B.STATUS <> 'C' WHERE B.CODTB3FAT is not null AND B.DATAEMISSAO  BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "') T1 " & _
                    "LEFT JOIN (SELECT A.FCE,d.nome,CASE WHEN max(a.adiantamento) is not null THEN max(a.adiantamento) else 0 end adiantamento,CASE WHEN max(B.databook) is not null THEN max(b.databook) else 0 end databook,b.dataentrega,SUM(A.PESO) AS PESO_TOTAL,SUM(A.TOTAL) AS VALOR_TOTAL FROM TBPEDIDOS AS A inner join tbFCE as b on a.fce = b.fce inner join tbFo as c on b.fce = c.fce inner join tbclifor as d " & _
                    "on c.codclifor = d.codclifor GROUP BY a.fce,d.nome,b.dataentrega) T2 ON T1.CODTB3FAT = T2.FCE " & _
                    "LEFT JOIN (SELECT B.CODTB3FAT AS FCE,SUM(B.PESOBRUTO) AS PESO,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05','2.2.25') AND B.STATUS <> 'C' WHERE B.CODTB3FAT is not null GROUP BY B.CODTB3FAT) T3 ON T1.CODTB3FAT = T3.FCE "
    
    rsFatPeriodo.Open sqlFatPeriodo, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsFatPeriodo.ActiveConnection = Nothing
    Set ReportProgramacao = report1
    
    ReportProgramacao.DiscardSavedData
    ReportProgramacao.Database.SetDataSource rsFatPeriodo
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsFatPeriodo.Close
    Set rsFatPeriodo = Nothing
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
