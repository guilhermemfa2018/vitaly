VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRApropriacao 
   Caption         =   "Apropriação"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRApropriacao.frx":0000
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
Attribute VB_Name = "FCRApropriacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim report1 As New CRApropriacao
    Dim rsOS As New ADODB.Recordset
    Dim SqlOS As String
    
    Dim crystalOS As New CRAXDRT.Application
    Dim ReportApropriacao As CRAXDRT.Report
    
    
Msgbox vDataFilter1
    
    rsOS.CursorLocation = adUseClient
    
    SqlOS = "SET LANGUAGE 'Português'"
    rsOS.Open SqlOS, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    SqlOS = "select a.chapa,c.NOME,b.idprogramacao,b.idos,b.idoperacao,b.nomecc,a.codigobarra,CONVERT (VARCHAR, a.dataent, 103) as DATA_BATIDA,DATEPART(YEAR,a.dataent)AS ANO,DATEPART(MONTH,a.dataent)AS MESNUM,RIGHT('0'+ CONVERT(VARCHAR,cast(DATEPART(MONTH,a.dataent) as varchar)),2) + ' - ' + DATENAME( MONTH, a.dataent) as MES,DATEPART(DAY,a.dataent)AS DIA,CONVERT (VARCHAR, a.horaent, 108) as ENT_OS,CONVERT (VARCHAR, a.HORASAI, 108) as SAI_OS,a.idparada,e.nmparada,g.NOME as nmCC,j.fce " & _
            "from tbOsMov as a left join tbMPItens as b on a.codigobarra = b.codigobarra inner join CORPORERM.dbo.PFUNC as c on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = c.CHAPA inner join CORPORERM.dbo.PPESSOA as d on c.CODPESSOA = d.CODIGO left join tbParadas as e on a.idparada = e.codigo left join tbMP as i on b.idprogramacao = i.idprogramacao inner join tbProjetos as j on i.codprojeto = j.codprojeto inner join CORPORERM.dbo.PFRATEIOFIXO as f on c.chapa = f.CHAPA inner join CORPORERM.dbo.GCCUSTO as g " & _
            "on f.CODCCUSTO = g.CODCCUSTO where b.idcc like '3000%' and a.dataent BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' AND  '" & Format(vDataFilter2, "yyyy/mm/dd") & "' or b.idcc like '7000.7103%' and a.dataent BETWEEN '" & Format(DTPicker1, "yyyy/mm/dd") & "' AND  '" & Format(DTPicker2, "yyyy/mm/dd") & "' order by ANO,MESNUM,NOME,DIA,ENT_OS"
    
    
    'SqlOS = "select a.chapa,c.NOME,b.idprogramacao,b.idos,b.idoperacao,b.nomecc,a.codigobarra,CONVERT (VARCHAR, a.dataent, 103) as DATA_BATIDA,DATEPART(YEAR,a.dataent)AS ANO,DATEPART(MONTH,a.dataent)AS MESNUM," & _
    '        "RIGHT('0'+ CONVERT(VARCHAR,cast(DATEPART(MONTH,a.dataent) as varchar)),2) + ' - ' + DATENAME( MONTH, a.dataent) as MES,DATEPART(DAY,a.dataent)AS DIA,CONVERT (VARCHAR, a.horaent, 108) as ENT_OS," & _
    '        "CONVERT (VARCHAR, a.HORASAI, 108) as SAI_OS,a.idparada,e.nmparada,g.NOME as nmCC from tbOsMov as a left join tbMPItens as b on a.codigobarra = b.codigobarra inner join " & vBancoTotvs & ".dbo.PFUNC as c on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = c.CHAPA " & _
    '        "inner join " & vBancoTotvs & ".dbo.PPESSOA as d on c.CODPESSOA = d.CODIGO left join tbParadas as e on a.idparada = e.codigo inner join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as f on c.chapa = f.CHAPA inner join " & vBancoTotvs & ".dbo.GCCUSTO as g on f.CODCCUSTO = g.CODCCUSTO " & _
    '        "where b.idcc like '3000%' or b.idcc like '7000.7103%' order by ANO,MESNUM,NOME,DIA,ENT_OS"
    rsOS.Open SqlOS, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsOS.ActiveConnection = Nothing
    Set ReportApropriacao = report1
    ReportApropriacao.DiscardSavedData
    ReportApropriacao.Database.SetDataSource rsOS
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
    FCRApropriacao.Hide
    Unload Me
    Set FCRApropriacao = Nothing
End Sub




