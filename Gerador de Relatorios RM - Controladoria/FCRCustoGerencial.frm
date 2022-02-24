VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRCustoGerencial 
   Caption         =   "Custo Gerencial"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRCustoGerencial.frx":0000
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
Attribute VB_Name = "FCRCustoGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRCustoGerencial

Private Sub Form_Load()
    Dim report1 As New CRCustoGerencial
    Dim rsCustoGerencial As New ADODB.Recordset
    Dim sqlCustoGerencial As String
    
    rsCustoGerencial.CursorLocation = adUseClient
    
    If vCustos <> "" Then
        sqlCustoGerencial = "SELECT SUBSTRING(I.CODTB5FAT,1,1) AS COD_CUSTOGER,I.DESCRICAO AS DES_CUSTOGER,SUBSTRING(H.CODTB5FAT,3,2) AS COD_SUBCENTRO1,H.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(G.CODTB5FAT,6,2) AS COD_SUBCENTRO2,G.DESCRICAO AS DES_SUBCENTRO," & _
                     "E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE,C.VALORUNITARIO AS PRECOUNIT,C.QUANTIDADEARECEBER*C.precounitario AS VALORTOTALITENS,DATEPART(YEAR,A.DATAEMISSAO) AS ANO," & _
                     "DATEPART(MONTH,A.DATAEMISSAO) AS MES,DATEPART(DAY,A.DATAEMISSAO) AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF,A.RECCREATEDON FROM TMOV AS A INNER JOIN FCFO AS B ON A.CODTMV IN(" & vMovs & ") and " & _
                     "A.CODCOLIGADA = B.CODCOLIGADA AND A.CODCFO = B.CODCFO AND A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV " & _
                     "INNER JOIN TCPG AS D ON A.CODCPG = D.CODCPG AND A.CODCOLIGADA=D.CODCOLIGADA INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA AND E.CODTB5FAT like '" & vCustos & "%' INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN TTB5 AS G ON E.CODTB5FAT = G.CODTB5FAT AND A.CODCOLIGADA=G.CODCOLIGADA " & _
                     "INNER JOIN TTB5 AS H ON H.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,4)+'.00.00' AND A.CODCOLIGADA=H.CODCOLIGADA INNER JOIN TTB5 AS I ON I.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,1)+'.00.00.00' AND A.CODCOLIGADA=I.CODCOLIGADA where a.CODCOLIGADA = " & vColigada & " " & _
                     "GROUP BY SUBSTRING(I.CODTB5FAT,1,1),I.DESCRICAO,SUBSTRING(H.CODTB5FAT,3,2),H.DESCRICAO,SUBSTRING(G.CODTB5FAT,6,2),G.DESCRICAO,E.CODIGOPRD,E.NOMEFANTASIA,E.CODUNDCONTROLE,C.QUANTIDADETOTAL,C.VALORUNITARIO,C.QUANTIDADEARECEBER*C.precounitario,DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO),DATEPART(DAY,A.DATAEMISSAO),A.CODTMV,A.IDMOV,A.NUMEROMOV,A.RECCREATEDON " & _
                     "ORDER BY ANO,MES,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
    ElseIf vProduto <> "" Then
        sqlCustoGerencial = "SELECT SUBSTRING(I.CODTB5FAT,1,1) AS COD_CUSTOGER,I.DESCRICAO AS DES_CUSTOGER,SUBSTRING(H.CODTB5FAT,3,2) AS COD_SUBCENTRO1,H.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(G.CODTB5FAT,6,2) AS COD_SUBCENTRO2,G.DESCRICAO AS DES_SUBCENTRO," & _
                     "E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE,C.VALORUNITARIO AS PRECOUNIT,C.QUANTIDADEARECEBER*C.precounitario AS VALORTOTALITENS,DATEPART(YEAR,A.DATAEMISSAO) AS ANO," & _
                     "DATEPART(MONTH,A.DATAEMISSAO) AS MES,DATEPART(DAY,A.DATAEMISSAO) AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF,A.RECCREATEDON FROM TMOV AS A INNER JOIN FCFO AS B ON A.CODTMV IN(" & vMovs & ") and " & _
                     "A.CODCOLIGADA = B.CODCOLIGADA AND A.CODCFO = B.CODCFO AND A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV " & _
                     "INNER JOIN TCPG AS D ON A.CODCPG = D.CODCPG AND A.CODCOLIGADA=D.CODCOLIGADA INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA AND E.NOMEFANTASIA LIKE '" & vProduto & "%' INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN TTB5 AS G ON E.CODTB5FAT = G.CODTB5FAT AND A.CODCOLIGADA=G.CODCOLIGADA " & _
                     "INNER JOIN TTB5 AS H ON H.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,4)+'.00.00' AND A.CODCOLIGADA=H.CODCOLIGADA INNER JOIN TTB5 AS I ON I.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,1)+'.00.00.00' AND A.CODCOLIGADA=I.CODCOLIGADA where a.CODCOLIGADA = " & vColigada & " " & _
                     "GROUP BY SUBSTRING(I.CODTB5FAT,1,1),I.DESCRICAO,SUBSTRING(H.CODTB5FAT,3,2),H.DESCRICAO,SUBSTRING(G.CODTB5FAT,6,2),G.DESCRICAO,E.CODIGOPRD,E.NOMEFANTASIA,E.CODUNDCONTROLE,C.QUANTIDADETOTAL,C.VALORUNITARIO,C.QUANTIDADEARECEBER*C.precounitario,DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO),DATEPART(DAY,A.DATAEMISSAO),A.CODTMV,A.IDMOV,A.NUMEROMOV,A.RECCREATEDON " & _
                     "ORDER BY ANO,MES,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
    Else
        sqlCustoGerencial = "SELECT SUBSTRING(I.CODTB5FAT,1,1) AS COD_CUSTOGER,I.DESCRICAO AS DES_CUSTOGER,SUBSTRING(H.CODTB5FAT,3,2) AS COD_SUBCENTRO1,H.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(G.CODTB5FAT,6,2) AS COD_SUBCENTRO2,G.DESCRICAO AS DES_SUBCENTRO," & _
                     "E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE,C.VALORUNITARIO AS PRECOUNIT,C.QUANTIDADEARECEBER*C.precounitario AS VALORTOTALITENS,DATEPART(YEAR,A.DATAEMISSAO) AS ANO," & _
                     "DATEPART(MONTH,A.DATAEMISSAO) AS MES,DATEPART(DAY,A.DATAEMISSAO) AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF,A.RECCREATEDON FROM TMOV AS A INNER JOIN FCFO AS B ON A.CODTMV IN(" & vMovs & ") and " & _
                     "A.CODCOLIGADA = B.CODCOLIGADA AND A.CODCFO = B.CODCFO AND A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV " & _
                     "INNER JOIN TCPG AS D ON A.CODCPG = D.CODCPG AND A.CODCOLIGADA=D.CODCOLIGADA INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN TTB5 AS G ON E.CODTB5FAT = G.CODTB5FAT AND A.CODCOLIGADA=G.CODCOLIGADA " & _
                     "INNER JOIN TTB5 AS H ON H.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,4)+'.00.00' AND A.CODCOLIGADA=H.CODCOLIGADA INNER JOIN TTB5 AS I ON I.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,1)+'.00.00.00' AND A.CODCOLIGADA=I.CODCOLIGADA where a.CODCOLIGADA = " & vColigada & " " & _
                     "GROUP BY SUBSTRING(I.CODTB5FAT,1,1),I.DESCRICAO,SUBSTRING(H.CODTB5FAT,3,2),H.DESCRICAO,SUBSTRING(G.CODTB5FAT,6,2),G.DESCRICAO,E.CODIGOPRD,E.NOMEFANTASIA,E.CODUNDCONTROLE,C.QUANTIDADETOTAL,C.VALORUNITARIO,C.QUANTIDADEARECEBER*C.precounitario,DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO),DATEPART(DAY,A.DATAEMISSAO),A.CODTMV,A.IDMOV,A.NUMEROMOV,A.RECCREATEDON " & _
                     "ORDER BY ANO,MES,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
    End If
    
    'sqlCustoGerencial = "SELECT SUBSTRING(I.CODTB5FAT,1,1) AS COD_CUSTOGER,I.DESCRICAO AS DES_CUSTOGER,SUBSTRING(H.CODTB5FAT,3,2) AS COD_SUBCENTRO1,H.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(G.CODTB5FAT,6,2) AS COD_SUBCENTRO2,G.DESCRICAO AS DES_SUBCENTRO," & _
    '                 "E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE,C.PRECOUNITARIO AS PRECOUNIT,C.QUANTIDADEARECEBER*C.PRECOUNITARIO AS VALORTOTALITENS,DATEPART(YEAR,A.DATAEMISSAO) AS ANO," & _
    '                 "DATEPART(MONTH,A.DATAEMISSAO) AS MES,DATEPART(DAY,A.DATAEMISSAO) AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF FROM TMOV AS A INNER JOIN FCFO AS B ON A.CODTMV IN('1.2.07','1.2.08','1.2.12','1.2.04','1.2.13','1.2.14') and " & _
    '                 "A.CODCOLIGADA = B.CODCOLIGADA AND A.CODCFO = B.CODCFO AND A.DATAEMISSAO BETWEEN '2013/01/01' and '2013/12/30' INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV " & _
    '                 "INNER JOIN TCPG AS D ON A.CODCPG = D.CODCPG INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV INNER JOIN TTB5 AS G ON E.CODTB5FAT = G.CODTB5FAT " & _
    '                 "INNER JOIN TTB5 AS H ON H.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,4)+'.00.00' INNER JOIN TTB5 AS I ON I.CODTB5FAT = SUBSTRING(G.CODTB5FAT,1,1)+'.00.00.00' ORDER BY ANO,MES,DIA,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsCustoGerencial.Open sqlCustoGerencial, cnBanco, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsCustoGerencial.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsCustoGerencial
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (118)
    Screen.MousePointer = vbDefault
    
    rsCustoGerencial.Close
    Set rsCustoGerencial = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRCustoGerencial.Hide
    Unload Me
    Set FCRCustoGerencial = Nothing
    Form1.MousePointer = 0
    Form1.Label1.Visible = False
End Sub



