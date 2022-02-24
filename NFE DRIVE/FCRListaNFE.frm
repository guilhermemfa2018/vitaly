VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form FCRListaNFE 
   Caption         =   "NFs não lançadas do RM"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "FCRListaNFE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
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
Attribute VB_Name = "FCRListaNFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRListaNFE

Private Sub Form_Load()
    Dim report1 As New CRListaNFE
    Dim rsOS As New ADODB.Recordset
    Dim SqlOS As String
    
    Dim crystalOS As New CRAXDRT.Application
    Dim ReportApropriacao As CRAXDRT.Report
    
    rsOS.CursorLocation = adUseClient
    
'    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    report1.Sections("PageHeaderSection1").ReportObjects("Text4").SetText frmimportarnfe.DTPicker1.Value
    report1.Sections("PageHeaderSection1").ReportObjects("Text17").SetText frmimportarnfe.DTPicker2.Value
    
    SqlOS = "SET LANGUAGE 'Português'"
    rsOS.Open SqlOS, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    SqlOS = ""
    SqlOS = SqlOS & "SELECT " & vbCrLf
    SqlOS = SqlOS & "     RIGHT(REPLICATE('0',6) + CONVERT(VARCHAR,A.NFE),6) AS NFE, " & vbCrLf
    SqlOS = SqlOS & "     A.SERIE, " & vbCrLf
    SqlOS = SqlOS & "     A.CNPJ, " & vbCrLf
    SqlOS = SqlOS & "     A.FORNECEDOR, " & vbCrLf
    SqlOS = SqlOS & "     CONVERT(VARCHAR, A.DTEMISSAO, 103) AS DTEMISSAO, " & vbCrLf
    SqlOS = SqlOS & "     CONVERT(VARCHAR, A.DTENTRADA, 103) AS DTENTRADA, " & vbCrLf
    SqlOS = SqlOS & "     A.VALORNF, " & vbCrLf
    SqlOS = SqlOS & "     B.CODCOLIGADA, " & vbCrLf
    SqlOS = SqlOS & "     B.CODFILIAL, " & vbCrLf
    SqlOS = SqlOS & "     B.CODTMV, " & vbCrLf
    SqlOS = SqlOS & "     B.NUMEROMOV, " & vbCrLf
    SqlOS = SqlOS & "     B.SERIE, " & vbCrLf
    SqlOS = SqlOS & "     C.CAMINHOLOGO " & vbCrLf
    SqlOS = SqlOS & "FROM TBNFE AS A " & vbCrLf
    SqlOS = SqlOS & "LEFT JOIN " & vBancoTotvs & ".DBO.TMOV AS B ON " & vbCrLf
    SqlOS = SqlOS & "     A.CODCOLIGADA = B.CODCOLIGADA AND " & vbCrLf
    SqlOS = SqlOS & "     right(replicate('0',9) + convert(VARCHAR,a.nfe),9) = right(replicate('0',9) + convert(VARCHAR,b.NUMEROMOV),9) COLLATE SQL_Latin1_General_CP1_CI_AS AND " & vbCrLf
    SqlOS = SqlOS & "     B.CODTMV IN('1.2.01','1.2.11','1.2.12','1.2.14','1.2.15','1.2.23','1.2.07','1.2.08','1.2.04','1.2.06','1.2.09','1.2.18','1.2.17','1.2.22') " & vbCrLf
    SqlOS = SqlOS & "LEFT JOIN TBLOGOCOLIGADA AS C ON " & vbCrLf
    SqlOS = SqlOS & "     A.CODCOLIGADA = C.CODCOLIGADA " & vbCrLf
    SqlOS = SqlOS & "WHERE " & vbCrLf
    SqlOS = SqlOS & "     B.NUMEROMOV IS NULL AND " & vbCrLf
    SqlOS = SqlOS & "     A.CODCOLIGADA = " & Val(Mid$(frmimportarnfe.Combo1.Text, 1, 6)) & "  AND " & vbCrLf
    SqlOS = SqlOS & "     A.NFE <> '' AND " & vbCrLf
    SqlOS = SqlOS & "     A.DTEMISSAO BETWEEN  '" & Format(vDataFilter1, "dd/mm/yyyy") & "' AND  '" & Format(vDataFilter2, "dd/mm/yyyy") & "'  " & vbCrLf
    SqlOS = SqlOS & "GROUP BY A.NFE,A.SERIE,A.CNPJ,A.FORNECEDOR,A.DTEMISSAO,A.DTENTRADA,A.VALORNF,B.CODCOLIGADA,B.CODFILIAL,B.CODTMV,B.NUMEROMOV,B.SERIE,C.CAMINHOLOGO " & vbCrLf
    SqlOS = SqlOS & "ORDER BY A.DTEMISSAO"
    
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
    Set rsOS = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRListaNFE.Hide
    Unload Me
    Set FCRListaNFE = Nothing
End Sub
