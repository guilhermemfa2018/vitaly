VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRFatFCESint 
   Caption         =   "Faturamento por FCE - Sintético"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRFatFCESint.frx":0000
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
Attribute VB_Name = "FCRFatFCESint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRFatFCESint

Private Sub Form_Load()
    Dim report1 As New CRFatFCESint
    Dim rsFatFCE As New ADODB.Recordset
    Dim sqlFatFCE As String
    Dim SomaPesoVendas As Double
    Dim SomaValorVendas As Double
    
    Dim rsTbTemp As New ADODB.Recordset
    Dim sqlTbTemp As String
    
    rsFatFCE.CursorLocation = adUseClient
    
'    sqlFatFCE = "SELECT T1.DESCRICAO,T1.CODTB3FAT,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO,T1.VALOR_LIQUIDO,T1.DTCRIACAO,T2.VALOR_ORIGINAL,T2.VALOR_BAIXADO,T2.VALOR_RECEBER,T3.PESO,T3.VALOR_TOTAL FROM " & _
'                "    (select  " & _
'                "        MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO,SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO " & _
'                "    FROM CORPORERM.dbo.TTB3 as a " & _
'                "    LEFT JOIN CORPORERM.dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' " & _
'                "    GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) " & _
'                "    ) T1 " & _
'                "LEFT JOIN  " & _
'                "    (SELECT " & _
'                "        B.CODTB3FAT,SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO) AS VALOR_BAIXADO, SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER  " & _
'                "    FROM CORPORERM.dbo.TTB3 as a " & _
'                "    LEFT JOIN CORPORERM.dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' " & _
'                "    INNER JOIN CORPORERM.dbo.FLAN AS C ON B.IDMOV = C.IDMOV " & _
'                "    GROUP BY B.CODTB3FAT " & _
'                "    ) T2 " & _
'                "ON T1.CODTB3FAT = T2.CODTB3FAT " & _
'                "LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from ZEUS.DBO.tbPedidos AS A group by a.fce ) T3 " & _
'                "ON T1.CODTB3FAT = T3.FCE where T2.VALOR_RECEBER > 0 or T2.VALOR_RECEBER is null ORDER BY T1.CODTB3FAT "
    
    If IsNull(frmPrintRels.DTPicker1.Value) Then
    sqlFatFCE = "SELECT T1.CODTB3FAT,T1.DESCRICAO,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO-isnull(T4.DEVOLVIDO,0) VALOR_BRUTO,T1.VALOR_LIQUIDO-isnull(T4.DEVOLVIDO,0) VALOR_LIQUIDO,T1.DTCRIACAO,(T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL,T2.VALOR_BAIXADO as VALOR_BAIXADO,((T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-isnull(T5.ADIANTADO,0)-isnull(T6.CANCELADO,0) AS VALOR_RECEBER, " & _
                "T3.PESO,T3.VALOR_TOTAL FROM (select MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO,SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50)) T1 LEFT JOIN (SELECT B.CODTB3FAT, " & _
                "SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY B.CODTB3FAT) T2 ON T1.CODTB3FAT = T2.CODTB3FAT LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from " & sDatabaseName & ".DBO.tbPedidos AS A group by a.fce) T3 ON T1.CODTB3FAT = T3.FCE LEFT JOIN (select B.CODTB3FAT, " & _
                "sum(b.VALORLIQUIDO) as DEVOLVIDO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT WHERE B.CODTMV in ('1.2.15','1.2.17') and B.STATUS = 'F' group by B.CODTB3FAT) T4 ON T1.CODTB3FAT = T4.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as ADIANTADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.25') group by B.CODTB3FAT) T5 ON T1.CODTB3FAT = T5.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as CANCELADO " & _
                "FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT WHERE B.STATUS = 'C' and B.CODTMV in ('2.2.01','2.2.05','1.2.15','1.2.17') group by b.CODTB3FAT) T6 ON T1.CODTB3FAT = T6.CODTB3FAT where T2.VALOR_RECEBER > 0 or T2.VALOR_RECEBER is null ORDER BY T1.CODTB3FAT"
    Else
    sqlFatFCE = "SELECT T1.CODTB3FAT,T1.DESCRICAO,T1.PESO_LIQUIDO,T1.PESO_BRUTO,T1.VALOR_BRUTO-isnull(T4.DEVOLVIDO,0) VALOR_BRUTO,T1.VALOR_LIQUIDO-isnull(T4.DEVOLVIDO,0) VALOR_LIQUIDO,T1.DTCRIACAO,(T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2) AS VALOR_ORIGINAL,T2.VALOR_BAIXADO as VALOR_BAIXADO,((T2.VALOR_ORIGINAL+isnull(T4.DEVOLVIDO,0)+isnull(T5.ADIANTADO,0)+isnull(T6.CANCELADO,0))-(isnull(T4.DEVOLVIDO,0)*2))-T2.VALOR_BAIXADO-isnull(T5.ADIANTADO,0)-isnull(T6.CANCELADO,0) AS VALOR_RECEBER, " & _
                "T3.PESO,T3.VALOR_TOTAL FROM (select MAX(B.IDMOV) AS IDMOV,a.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50) AS DESCRICAO,SUM(B.PESOLIQUIDO) AS PESO_LIQUIDO,SUM(B.PESOBRUTO) AS PESO_BRUTO,SUM(B.VALORBRUTO) AS VALOR_LIQUIDO,SUM(B.VALORLIQUIDO) AS VALOR_BRUTO,MAX(CONVERT (VARCHAR, A.RECCREATEDON, 103)) as DTCRIACAO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' where A.RECCREATEDON BETWEEN '" & Format(frmPrintRels.DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(frmPrintRels.DTPicker2.Value, "yyyy/mm/dd") & "' GROUP BY A.CODTB3FAT,SUBSTRING(a.DESCRICAO,1,50)) T1 LEFT JOIN (SELECT B.CODTB3FAT, " & _
                "SUM(C.VALORORIGINAL) AS VALOR_ORIGINAL,SUM(C.VALORBAIXADO+C.VALORADIANTAMENTO) AS VALOR_BAIXADO,SUM(C.VALORORIGINAL-C.VALORBAIXADO) as VALOR_RECEBER FROM " & vBancoTotvs & ".dbo.TMOV AS B INNER JOIN " & vBancoTotvs & ".dbo.FLAN AS C ON B.IDMOV = C.IDMOV AND B.CODTMV in ('2.2.01','2.2.05') AND B.STATUS <> 'C' GROUP BY B.CODTB3FAT) T2 ON T1.CODTB3FAT = T2.CODTB3FAT LEFT JOIN (select a.fce AS fce,SUM(a.peso) AS PESO,SUM(a.total) AS VALOR_TOTAL from " & sDatabaseName & ".DBO.tbPedidos AS A group by a.fce) T3 ON T1.CODTB3FAT = T3.FCE LEFT JOIN (select B.CODTB3FAT, " & _
                "sum(b.VALORLIQUIDO) as DEVOLVIDO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT WHERE B.CODTMV in ('1.2.15','1.2.17') and B.STATUS = 'F' group by B.CODTB3FAT) T4 ON T1.CODTB3FAT = T4.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as ADIANTADO FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT AND B.CODTMV in ('2.2.25') group by B.CODTB3FAT) T5 ON T1.CODTB3FAT = T5.CODTB3FAT LEFT JOIN (select B.CODTB3FAT,sum(b.VALORLIQUIDO) as CANCELADO " & _
                "FROM " & vBancoTotvs & ".dbo.TTB3 as a LEFT JOIN " & vBancoTotvs & ".dbo.TMOV AS B ON A.CODTB3FAT = B.CODTB3FAT WHERE B.STATUS = 'C' and B.CODTMV in ('2.2.01','2.2.05','1.2.15','1.2.17') group by b.CODTB3FAT) T6 ON T1.CODTB3FAT = T6.CODTB3FAT where T2.VALOR_RECEBER > 0 or T2.VALOR_RECEBER is null ORDER BY T1.CODTB3FAT"
    End If
    
' A.RECCREATEDON BETWEEN '" & Format(frmPrintRels.DTPicker1.Value, "yyyy/mm/dd") & "' AND  '" & Format(frmPrintRels.DTPicker2.Value, "yyyy/mm/dd") & "'
'and T1.DTCRIACAO BETWEEN '2015/01/01' AND  '2015/05/12'
    
    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsFatFCE.Open sqlFatFCE, cnBanco, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsFatFCE.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsFatFCE
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (140)
    Screen.MousePointer = vbDefault
    
    rsFatFCE.Close
    Set rsFatFCE = Nothing
    Exit Sub

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'    FCRFatFCE.Hide
    Set FCRFatFCESint = Nothing
    Unload Me
    'Form1.MousePointer = 0
    'Form1.Label1.Visible = False
End Sub


