VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCREstoque 
   Caption         =   "Custo Gerencial"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCREstoque.frx":0000
   LinkTopic       =   "Form2"
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
Attribute VB_Name = "FCREstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CREstoque

Private Sub Form_Load()
    Dim report1 As New CREstoque
    Dim rsEstoque As New ADODB.Recordset
    Dim sqlEstoque As String
    
    rsEstoque.CursorLocation = adUseClient
    
'    sqlEstoque = "select T1.IDPRD,T1.CODPRODUTO AS CODPRODUTO,T1.NOMEFANTASIA AS PRODUTO,T1.SALDOFISICO2 AS QTDE,T1.UND,cast(T1.PRECO_UNITARIO as decimal(10, 2)) AS PRECOUNIT,T1.SALDOFISICO2*T1.PRECO_UNITARIO AS VALORTOTALITENS,T1.COD_CUSTOGER,T1.DES_CUSTOGER,T1.COD_SUBCENTRO1,T1.DES_SUBCENTRO1,T1.COD_SUBCENTRO2,T1.DES_SUBCENTRO,max(T1.CODLOC) AS CODLOC from " & _
'                 "(select a.IDPRD AS IDPRD,b.CODIGOPRD AS CODPRODUTO,b.NOMEFANTASIA,a.SALDOFISICO2,b.CODUNDCONTROLE AS UND,(select top 1 CONVERT (VARCHAR, c.PRECOUNITARIO, 103) as DATAENTREGA1 from TITMMOV AS C where c.IDPRD = b.IDPRD and c.DATAENTREGA is not null order by c.DATAENTREGA desc) as PRECO_UNITARIO,SUBSTRING(I.CODTB5FAT,1,1) AS COD_CUSTOGER,I.DESCRICAO AS DES_CUSTOGER,SUBSTRING(H.CODTB5FAT,3,2) AS COD_SUBCENTRO1,H.DESCRICAO AS DES_SUBCENTRO1, " & _
'                 "SUBSTRING(G.CODTB5FAT,6,2) AS COD_SUBCENTRO2,G.DESCRICAO AS DES_SUBCENTRO,a.CODLOC from TPRDLOC as a inner join TPRD as b on a.IDPRD = b.IDPRD and a.CODCOLIGADA = b.CODCOLIGADA LEFT JOIN TTB5 AS G ON B.CODTB5FAT = G.CODTB5FAT LEFT JOIN TTB5 AS H ON H.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,4)+'.00.00' LEFT JOIN TTB5 AS I ON I.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,1)+'.00.00.00' where a.SALDOFISICO2 > 0) T1 " & _
'                 "where T1.SALDOFISICO2 > 0 and T1.CODLOC IN(" & vMovs & ") GROUP BY T1.IDPRD,T1.CODPRODUTO,T1.NOMEFANTASIA,T1.UND,T1.PRECO_UNITARIO,T1.SALDOFISICO2,T1.COD_CUSTOGER,T1.DES_CUSTOGER,T1.COD_SUBCENTRO1,T1.DES_SUBCENTRO1,T1.COD_SUBCENTRO2,T1.DES_SUBCENTRO ORDER BY T1.IDPRD"
    
    sqlEstoque = "select T1.IDPRD,T1.CODPRODUTO AS CODPRODUTO,T1.NOMEFANTASIA AS PRODUTO,T1.SALDOFISICO2 AS QTDE,T1.UND,cast(T1.PRECO_UNITARIO as decimal(10, 2)) AS PRECOUNIT,T1.SALDOFISICO2*T1.PRECO_UNITARIO AS VALORTOTALITENS,T1.COD_CUSTOGER,T1.DES_CUSTOGER,T1.COD_SUBCENTRO1,T1.DES_SUBCENTRO1,T1.COD_SUBCENTRO2,T1.DES_SUBCENTRO,max(T1.CODLOC) AS CODLOC,T1.RECMODIFIEDBY,t1.RECMODIFIEDON from " & _
                 "(select a.IDPRD AS IDPRD,b.CODIGOPRD AS CODPRODUTO,b.NOMEFANTASIA,a.SALDOFISICO2,b.CODUNDCONTROLE AS UND,(select top 1 CONVERT (VARCHAR, c.PRECOUNITARIO, 103) as DATAENTREGA1 from TITMMOV AS C inner join TMOV as d on c.CODCOLIGADA = d.CODCOLIGADA and c.IDMOV = d.IDMOV where c.IDPRD = b.IDPRD and c.DATAENTREGA is not null and d.CODTMV like '1.2%' order by c.DATAEMISSAO desc) as PRECO_UNITARIO,SUBSTRING(I.CODTB5FAT,1,1) AS COD_CUSTOGER,I.DESCRICAO AS DES_CUSTOGER,SUBSTRING(H.CODTB5FAT,3,2) AS COD_SUBCENTRO1,H.DESCRICAO AS DES_SUBCENTRO1, " & _
                 "SUBSTRING(G.CODTB5FAT,6,2) AS COD_SUBCENTRO2,G.DESCRICAO AS DES_SUBCENTRO,a.CODLOC,a.RECMODIFIEDBY,CONVERT (VARCHAR, a.RECMODIFIEDON, 103) AS RECMODIFIEDON,A.CODCOLIGADA from TPRDLOC as a inner join TPRD as b on a.IDPRD = b.IDPRD and a.CODCOLIGADA = b.CODCOLIGADA LEFT JOIN TTB5 AS G ON B.CODTB5FAT = G.CODTB5FAT LEFT JOIN TTB5 AS H ON H.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,4)+'.00.00' LEFT JOIN TTB5 AS I ON I.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,1)+'.00.00.00' where a.SALDOFISICO2 > 0) T1 WHERE " & _
                 "T1.SALDOFISICO2 > 0 and T1.CODLOC IN(" & vMovs & ") AND T1.CODCOLIGADA = " & vColigada & " GROUP BY T1.IDPRD,T1.CODPRODUTO,T1.NOMEFANTASIA,T1.UND,T1.PRECO_UNITARIO,T1.SALDOFISICO2,T1.COD_CUSTOGER,T1.DES_CUSTOGER,T1.COD_SUBCENTRO1,T1.DES_SUBCENTRO1,T1.COD_SUBCENTRO2,T1.DES_SUBCENTRO,T1.RECMODIFIEDBY,T1.RECMODIFIEDON ORDER BY T1.IDPRD"
    
    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsEstoque.Open sqlEstoque, cnBanco, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsEstoque.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsEstoque
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (118)
    Screen.MousePointer = vbDefault
    
    rsEstoque.Close
    Set rsEstoque = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCREstoque.Hide
    Unload Me
    Set FCREstoque = Nothing
    Form1.MousePointer = 0
    Form1.Label1.Visible = False
End Sub


