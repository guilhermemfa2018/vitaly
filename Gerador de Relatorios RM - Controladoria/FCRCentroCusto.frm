VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "crviewer.dll"
Begin VB.Form FCRCentroCusto 
   Caption         =   "Centro de Custo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRCentroCusto.frx":0000
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
Attribute VB_Name = "FCRCentroCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRCentroCusto
Private Sub Form_Load()
    'On Error Resume Next
    Dim report1 As New CRCentroCusto
    Dim rsCentroCusto As New ADODB.Recordset
    Dim sqlCentroCusto As String
    
    Dim rsTbTemp As New ADODB.Recordset
    Dim sqlTbTemp As String
    
    
    rsCentroCusto.CursorLocation = adUseClient
    
    
'    If vFCECC <> "" Then
'        sqlCentroCusto = "SELECT G.CODCCUSTO,L.NOME AS SUBCENTRO1,M.NOME AS SUBCENTRO2,G.CODREDUZIDO AS CCUSTO,G.NOME,C.IDPRD,E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE," & _
'                     "C.PRECOUNITARIO AS PRECOUNIT,(C.QUANTIDADEARECEBER*C.PRECOUNITARIO) AS VALORTOTALITENS,case WHEN A.CODTMV = '2.2.22' OR A.CODTMV = '1.2.10' then N.PRECOUNIT else C.PRECOUNITARIO END AS PRECOUNITENT,(C.QUANTIDADEARECEBER*N.PRECOUNIT) AS VALORTOTALREAL," & _
'                     "DATEPART(YEAR,A.DATAEMISSAO)AS ANO,DATEPART(MONTH,A.DATAEMISSAO)AS MES,DATEPART(DAY,A.DATAEMISSAO)AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF,A.CAMPOLIVRE3 AS FCE_MANUAL,A.CODTB3FAT AS FCE_CC,H.DESCRICAO," & _
'                     "SUBSTRING(K.CODTB5FAT,1,1) AS COD_CUSTOGER,K.DESCRICAO AS DES_CUSTOGER,SUBSTRING(J.CODTB5FAT,3,2) AS COD_SUBCENTRO1,J.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(I.CODTB5FAT,6,2) AS COD_SUBCENTRO2,I.DESCRICAO AS DES_SUBCENTRO,A.RECCREATEDON " & _
'                     "FROM TMOV AS A INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV AND A.CODTMV IN(" & vMovs & ") AND A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' AND A.CODTB3FAT LIKE '" & vFCECC & "%' OR " & _
'                     "A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV AND A.CODTMV IN(" & vMovs & ") AND A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' AND A.CAMPOLIVRE3 LIKE '" & vFCECC & "%' " & _
'                     "INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN GCCUSTO AS G " & _
'                     "ON A.CODCCUSTO = G.CODCCUSTO AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN GCCUSTO AS L ON SUBSTRING(A.CODCCUSTO,1,4) = L.CODCCUSTO AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN GCCUSTO AS M ON SUBSTRING(A.CODCCUSTO,1,7) = M.CODCCUSTO AND A.CODCOLIGADA=M.CODCOLIGADA LEFT JOIN TTB3 AS H ON A.CODTB3FAT = H.CODTB3FAT and H.CODCOLIGADA = " & vColigada & " LEFT JOIN TTB5 AS I " & _
'                     "ON E.CODTB5FAT = I.CODTB5FAT AND A.CODCOLIGADA=I.CODCOLIGADA LEFT JOIN TTB5 AS J ON J.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,4)+'.00.00' AND A.CODCOLIGADA=J.CODCOLIGADA LEFT JOIN TTB5 AS K ON K.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,1)+'.00.00.00' AND A.CODCOLIGADA=K.CODCOLIGADA LEFT JOIN VALORENTTEMP AS N ON E.IDPRD  = N.IDPRD " & _
'                     "where a.CODCOLIGADA = " & vColigada & " and (A.CODTMV <> '1.2.07' or SUBSTRING(K.CODTB5FAT,1,1) >1 and SUBSTRING(K.CODTB5FAT,1,1) < 10) " & _
'                     "group by G.CODCCUSTO,L.NOME,M.NOME,G.CODREDUZIDO,G.NOME,C.IDPRD,E.CODIGOPRD,E.NOMEFANTASIA,E.CODUNDCONTROLE,C.QUANTIDADETOTAL,C.PRECOUNITARIO,(C.QUANTIDADEARECEBER*C.PRECOUNITARIO),case WHEN A.CODTMV = '2.2.22' OR A.CODTMV = '1.2.10' then N.PRECOUNIT else C.PRECOUNITARIO END,(C.QUANTIDADEARECEBER*N.PRECOUNIT),DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO), " & _
'                     "DATEPART(DAY,A.DATAEMISSAO),A.CODTMV,A.IDMOV,A.NUMEROMOV,A.CAMPOLIVRE3,A.CODTB3FAT,H.DESCRICAO,SUBSTRING(K.CODTB5FAT,1,1),K.DESCRICAO,SUBSTRING(J.CODTB5FAT,3,2),J.DESCRICAO,SUBSTRING(I.CODTB5FAT,6,2),I.DESCRICAO,A.RECCREATEDON  " & _
'                     "ORDER BY ANO,MES,DIA,SUBCENTRO1,SUBCENTRO2,CCUSTO,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
'    ElseIf vCustos <> "" Then
'        sqlCentroCusto = "SELECT G.CODCCUSTO,L.NOME AS SUBCENTRO1,M.NOME AS SUBCENTRO2,G.CODREDUZIDO AS CCUSTO,G.NOME,C.IDPRD,E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE," & _
'                     "C.PRECOUNITARIO AS PRECOUNIT,(C.QUANTIDADEARECEBER*C.PRECOUNITARIO) AS VALORTOTALITENS,case WHEN A.CODTMV = '2.2.22' OR A.CODTMV = '1.2.10' then N.PRECOUNIT else C.PRECOUNITARIO END AS PRECOUNITENT,(C.QUANTIDADEARECEBER*N.PRECOUNIT) AS VALORTOTALREAL," & _
'                     "DATEPART(YEAR,A.DATAEMISSAO)AS ANO,DATEPART(MONTH,A.DATAEMISSAO)AS MES,DATEPART(DAY,A.DATAEMISSAO)AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF,A.CAMPOLIVRE3 AS FCE_MANUAL,A.CODTB3FAT AS FCE_CC,H.DESCRICAO," & _
'                     "SUBSTRING(K.CODTB5FAT,1,1) AS COD_CUSTOGER,K.DESCRICAO AS DES_CUSTOGER,SUBSTRING(J.CODTB5FAT,3,2) AS COD_SUBCENTRO1,J.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(I.CODTB5FAT,6,2) AS COD_SUBCENTRO2,I.DESCRICAO AS DES_SUBCENTRO,A.RECCREATEDON " & _
'                     "FROM TMOV AS A INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV AND A.CODTMV IN(" & vMovs & ") AND " & _
'                     "A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN GCCUSTO AS G " & _
'                     "ON A.CODCCUSTO = G.CODCCUSTO AND A.CODCOLIGADA=G.CODCOLIGADA AND G.CODREDUZIDO like '%" & vCustos & "%' INNER JOIN GCCUSTO AS L ON SUBSTRING(A.CODCCUSTO,1,4) = L.CODCCUSTO AND A.CODCOLIGADA=L.CODCOLIGADA INNER JOIN GCCUSTO AS M ON SUBSTRING(A.CODCCUSTO,1,7) = M.CODCCUSTO AND A.CODCOLIGADA=M.CODCOLIGADA LEFT JOIN TTB3 AS H ON A.CODTB3FAT = H.CODTB3FAT and H.CODCOLIGADA = " & vColigada & " LEFT JOIN TTB5 AS I " & _
'                     "ON E.CODTB5FAT = I.CODTB5FAT AND A.CODCOLIGADA=I.CODCOLIGADA LEFT JOIN TTB5 AS J ON J.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,4)+'.00.00' AND A.CODCOLIGADA=K.CODCOLIGADA LEFT JOIN TTB5 AS K ON K.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,1)+'.00.00.00' AND A.CODCOLIGADA=K.CODCOLIGADA LEFT JOIN VALORENTTEMP AS N ON E.IDPRD  = N.IDPRD " & _
'                     "where a.CODCOLIGADA = " & vColigada & " and (A.CODTMV <> '1.2.07' or SUBSTRING(K.CODTB5FAT,1,1) >1 and SUBSTRING(K.CODTB5FAT,1,1) < 10) " & _
'                     "group by G.CODCCUSTO,L.NOME,M.NOME,G.CODREDUZIDO,G.NOME,C.IDPRD,E.CODIGOPRD,E.NOMEFANTASIA,E.CODUNDCONTROLE,C.QUANTIDADETOTAL,C.PRECOUNITARIO,(C.QUANTIDADEARECEBER*C.PRECOUNITARIO),case WHEN A.CODTMV = '2.2.22' OR A.CODTMV = '1.2.10' then N.PRECOUNIT else C.PRECOUNITARIO END,(C.QUANTIDADEARECEBER*N.PRECOUNIT),DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO), " & _
'                     "DATEPART(DAY,A.DATAEMISSAO),A.CODTMV,A.IDMOV,A.NUMEROMOV,A.CAMPOLIVRE3,A.CODTB3FAT,H.DESCRICAO,SUBSTRING(K.CODTB5FAT,1,1),K.DESCRICAO,SUBSTRING(J.CODTB5FAT,3,2),J.DESCRICAO,SUBSTRING(I.CODTB5FAT,6,2),I.DESCRICAO,A.RECCREATEDON  " & _
'                     "ORDER BY ANO,MES,DIA,SUBCENTRO1,SUBCENTRO2,CCUSTO,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
'    ElseIf vProduto <> "" Then
'        sqlCentroCusto = "SELECT G.CODCCUSTO,L.NOME AS SUBCENTRO1,M.NOME AS SUBCENTRO2,G.CODREDUZIDO AS CCUSTO,G.NOME,C.IDPRD,E.CODIGOPRD AS CODPRODUTO,E.NOMEFANTASIA AS PRODUTO,E.CODUNDCONTROLE AS UND,C.QUANTIDADETOTAL AS QTDE," & _
'                     "C.PRECOUNITARIO AS PRECOUNIT,(C.QUANTIDADEARECEBER*C.PRECOUNITARIO) AS VALORTOTALITENS,case WHEN A.CODTMV = '2.2.22' OR A.CODTMV = '1.2.10' then N.PRECOUNIT else C.PRECOUNITARIO END AS PRECOUNITENT,(C.QUANTIDADEARECEBER*N.PRECOUNIT) AS VALORTOTALREAL," & _
'                     "DATEPART(YEAR,A.DATAEMISSAO)AS ANO,DATEPART(MONTH,A.DATAEMISSAO)AS MES,DATEPART(DAY,A.DATAEMISSAO)AS DIA,A.CODTMV AS TPMOV,A.IDMOV AS MOVIMENTONUM,A.NUMEROMOV AS NUMNF,A.CAMPOLIVRE3 AS FCE_MANUAL,A.CODTB3FAT AS FCE_CC,H.DESCRICAO," & _
'                     "SUBSTRING(K.CODTB5FAT,1,1) AS COD_CUSTOGER,K.DESCRICAO AS DES_CUSTOGER,SUBSTRING(J.CODTB5FAT,3,2) AS COD_SUBCENTRO1,J.DESCRICAO AS DES_SUBCENTRO1,SUBSTRING(I.CODTB5FAT,6,2) AS COD_SUBCENTRO2,I.DESCRICAO AS DES_SUBCENTRO,A.RECCREATEDON " & _
'                     "FROM TMOV AS A INNER JOIN TITMMOV AS C ON A.CODCOLIGADA = C.CODCOLIGADA AND A.IDMOV = C.IDMOV AND A.CODTMV IN(" & vMovs & ") AND " & _
'                     "A.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' INNER JOIN TPRD AS E ON C.IDPRD = E.IDPRD AND A.CODCOLIGADA = E.CODCOLIGADA AND E.NOMEFANTASIA LIKE '" & vProduto & "%' INNER JOIN TITMMOVCOMPL AS F ON C.IDMOV = F.IDMOV AND C.NSEQITMMOV = F.NSEQITMMOV AND A.CODCOLIGADA=F.CODCOLIGADA INNER JOIN GCCUSTO AS G " & _
'                     "ON A.CODCCUSTO = G.CODCCUSTO AND A.CODCOLIGADA=G.CODCOLIGADA INNER JOIN GCCUSTO AS L ON SUBSTRING(A.CODCCUSTO,1,4) = L.CODCCUSTO AND A.CODCOLIGADA=L.CODCOLIGADA INNER JOIN GCCUSTO AS M ON SUBSTRING(A.CODCCUSTO,1,7) = M.CODCCUSTO AND A.CODCOLIGADA=M.CODCOLIGADA LEFT JOIN TTB3 AS H ON A.CODTB3FAT = H.CODTB3FAT AND H.CODCOLIGADA = " & vColigada & "  LEFT JOIN TTB5 AS I " & _
'                     "ON E.CODTB5FAT = I.CODTB5FAT AND A.CODCOLIGADA=I.CODCOLIGADA LEFT JOIN TTB5 AS J ON J.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,4)+'.00.00' AND A.CODCOLIGADA=J.CODCOLIGADA LEFT JOIN TTB5 AS K ON K.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,1)+'.00.00.00' AND A.CODCOLIGADA=K.CODCOLIGADA LEFT JOIN VALORENTTEMP AS N ON E.IDPRD  = N.IDPRD " & _
'                     "where a.CODCOLIGADA = " & vColigada & " and (A.CODTMV <> '1.2.07' or SUBSTRING(K.CODTB5FAT,1,1) >1 and SUBSTRING(K.CODTB5FAT,1,1) < 10) " & _
'                     "group by G.CODCCUSTO,L.NOME,M.NOME,G.CODREDUZIDO,G.NOME,C.IDPRD,E.CODIGOPRD,E.NOMEFANTASIA,E.CODUNDCONTROLE,C.QUANTIDADETOTAL,C.PRECOUNITARIO,(C.QUANTIDADEARECEBER*C.PRECOUNITARIO),case WHEN A.CODTMV = '2.2.22' OR A.CODTMV = '1.2.10' then N.PRECOUNIT else C.PRECOUNITARIO END,(C.QUANTIDADEARECEBER*N.PRECOUNIT),DATEPART(YEAR,A.DATAEMISSAO),DATEPART(MONTH,A.DATAEMISSAO), " & _
'                     "DATEPART(DAY,A.DATAEMISSAO),A.CODTMV,A.IDMOV,A.NUMEROMOV,A.CAMPOLIVRE3,A.CODTB3FAT,H.DESCRICAO,SUBSTRING(K.CODTB5FAT,1,1),K.DESCRICAO,SUBSTRING(J.CODTB5FAT,3,2),J.DESCRICAO,SUBSTRING(I.CODTB5FAT,6,2),I.DESCRICAO,A.RECCREATEDON  " & _
'                     "ORDER BY ANO,MES,DIA,SUBCENTRO1,SUBCENTRO2,CCUSTO,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"
'    Else
        sqlCentroCusto = ""
        sqlCentroCusto = sqlCentroCusto & "SELECT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CODCCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " IDPRD, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CODPRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " PRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " UND, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " QTDE, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " PRECOUNIT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " VALORTOTALITENS, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " PRECOUNITENT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " ANO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " MES, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " DIA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " TPMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " IDMOV_SAIDA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " NUMNF, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " FCE_MANUAL, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " FCE_CC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " COD_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " DES_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " COD_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " DES_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " COD_SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " DES_SUBCENTRO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " RECCREATEDON, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CASE WHEN [IPI] IS NULL THEN CAST(0 AS DECIMAL(7,2)) ELSE CAST([IPI] AS DECIMAL(7,2)) END AS IPI, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CASE WHEN [ICMS] IS NULL THEN CAST(0 AS DECIMAL(7,2)) ELSE CAST([ICMS] AS DECIMAL(7,2)) END AS ICMS, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CASE WHEN [PIS001] IS NULL THEN CAST(0 AS DECIMAL(7,2)) ELSE CAST([PIS001] AS DECIMAL(7,2)) END PIS001, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CASE WHEN [COF001] IS NULL THEN CAST(0 AS DECIMAL(7,2)) ELSE CAST([COF001] AS DECIMAL(7,2)) END COF001, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " IDMOV_ENTRADA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " NSEQITMMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " CODTB4FAT_ENT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " VALORDESC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " VALORDESP, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " VALORFRETECTRC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " RATEIOFRETE, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " VALOR_C_IMPOSTOS = " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     CASE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         WHEN CODTB4FAT_ENT = 0 THEN " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             CASE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 WHEN TPMOV = '2.2.22' OR TPMOV = '1.2.10' THEN " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     (PRECOUNITENT*QTDE) -  ISNULL(IPI, 0)  + ISNULL(ICMS,0) + ISNULL(PIS001,0) + ISNULL(COF001,0) + VALORDESC - VALORDESP - VALORFRETECTRC - RATEIOFRETE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 ELSE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     (PRECOUNIT*QTDE) " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             END " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ELSE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             CASE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 WHEN TPMOV = '2.2.22' OR TPMOV = '1.2.10' THEN " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     (PRECOUNITENT*QTDE) + ISNULL(ICMS,0) + ISNULL(PIS001,0) + ISNULL(COF001,0) + VALORDESC -  VALORDESP - VALORFRETECTRC - RATEIOFRETE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 ELSE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     (PRECOUNIT*QTDE) " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             END " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     END " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "FROM " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " ( " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     /*INICIO DA SEGUNDA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     /*Exibe os tributos dos movimentos*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     SELECT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CODCCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDPRD, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CODPRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.PRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.UND, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.QTDE, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.PRECOUNIT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.VALORTOTALITENS, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.PRECOUNITENT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.ANO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.MES, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DIA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.TPMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDMOV_SAIDA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.NUMNF, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.FCE_MANUAL, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.FCE_CC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.COD_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DES_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.COD_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DES_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.COD_SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DES_SUBCENTRO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.RECCREATEDON, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDMOV_ENTRADA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         TRIBUTOS.CODTRB AS COD_TRIBUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         CASE WHEN TRIBUTOS.ALIQUOTA IS NULL THEN 0.0000 ELSE (SEM_TRIBUTOS.PRECOUNITENT*TRIBUTOS.ALIQUOTA)/100 END AS ALIQUOTA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         MAX(TRIBUTOS.NSEQITMMOV) AS NSEQITMMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         CODTB4FAT_ENT = " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             ( " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 SELECT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     A.CODTB4FAT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 FROM TMOV AS A " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 WHERE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     SEM_TRIBUTOS.IDMOV_ENTRADA = A.IDMOV AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     A.CODCOLIGADA = SEM_TRIBUTOS.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             ), " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ISNULL(MAX(OUTRAS_DESP.VALORDESC),0) AS VALORDESC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ISNULL(MAX(OUTRAS_DESP.VALORDESP),0) AS VALORDESP, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ISNULL(MAX(OUTRAS_DESP.VALORFRETECTRC),0) AS VALORFRETECTRC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ISNULL(MAX(OUTRAS_DESP.RATEIOFRETE),0) AS RATEIOFRETE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     FROM " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ( " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             /*INICIO DA SEGUNDA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             /*Aqui busca os dados dados de entrada dos movimentos 2.2.22 e 1.2.10 nos tipos de movimento 1.2.07 e 1.2.23 */ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             SELECT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 CUSTO.*, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 PRECOUNITENT = " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     CASE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         /*DESCONSIDERA O VALOR REGISTRADO NO MOVIMENTO E BUSCA NAS TABELAS O VALOR DA ULTIMA COMPRA DO PRODUTO*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         WHEN CUSTO.TPMOV = '2.2.22' OR CUSTO.TPMOV = '1.2.10' THEN " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                             ( " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 SELECT TOP 1 " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     Y.PRECOUNITARIO AS PRECOUNIT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 FROM TMOV AS X " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 INNER JOIN TITMMOV AS Y ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     X.CODCOLIGADA = CUSTO.CODCOLIGADA AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     X.IDMOV = Y.IDMOV AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     X.CODTMV IN('1.2.07','1.2.23') AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     Y.IDPRD = CUSTO.IDPRD " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 ORDER BY X.DATASAIDA DESC " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                             ) " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                          /*PARA OS DEMAIS TIPOS DE MOVIMENTO, É CONSIDERADO O VALOR REGISTRADO NO MOVIMENTO*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         ELSE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                             CUSTO.PRECOUNIT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     END, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 IDMOV_ENTRADA = " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     CASE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         /*PEGA O IDMOV DA ULTIMA ENTRADA DO PRODUTO EM UM DOS TIPOS DE MOVIMENTOS ABAIXO*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         WHEN CUSTO.TPMOV = '2.2.22' OR CUSTO.TPMOV = '1.2.10' then " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                             ( " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 SELECT TOP 1 " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     Y.IDMOV " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 FROM TMOV AS X " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 INNER JOIN TITMMOV AS Y ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     X.CODCOLIGADA = CUSTO.CODCOLIGADA AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     X.IDMOV = Y.IDMOV AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     X.CODTMV IN('1.2.07','1.2.23') AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                     Y.IDPRD = CUSTO.IDPRD " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                                 ORDER BY X.DATASAIDA DESC " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                             ) " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         /*PARA OS DEMAIS MOVIMENTO O QUE VALE É O IDMOV DO MOVIMENTO*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         ELSE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                             CUSTO.IDMOV_SAIDA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     END " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             FROM ( " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     /*INICIO DA PRIMEIRA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     /*Busca os dados principais dos tipos de movimentos selecionados*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     SELECT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         G.CODCCUSTO AS CODCCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         L.NOME AS SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         M.NOME AS SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         G.CODREDUZIDO AS CCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         G.NOME AS NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.IDPRD AS IDPRD, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.CODIGOPRD AS CODPRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.NOMEFANTASIA AS PRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.CODUNDCONTROLE AS UND, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.QUANTIDADETOTAL AS QTDE, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.PRECOUNITARIO AS PRECOUNIT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         (C.QUANTIDADEARECEBER*C.PRECOUNITARIO) AS VALORTOTALITENS, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         DATEPART(YEAR,A.DATASAIDA)AS ANO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         DATEPART(MONTH,A.DATASAIDA)AS MES, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         DATEPART(DAY,A.DATASAIDA)AS DIA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.DATASAIDA AS DATAEMISSAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODTMV AS TPMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.IDMOV AS IDMOV_SAIDA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.NUMEROMOV AS NUMNF, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CAMPOLIVRE3 AS FCE_MANUAL, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODTB3FAT AS FCE_CC,H.DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(K.CODTB5FAT,1,1) AS COD_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         K.DESCRICAO AS DES_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(J.CODTB5FAT,3,2) AS COD_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         J.DESCRICAO AS DES_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(I.CODTB5FAT,6,2) AS COD_SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         I.DESCRICAO AS DES_SUBCENTRO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.RECCREATEDON AS RECCREATEDON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     FROM TMOV AS A " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     INNER JOIN TITMMOV AS C ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = C.CODCOLIGADA AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.IDMOV = C.IDMOV " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN TPRD AS E ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.IDPRD = E.IDPRD AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = E.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN GCCUSTO AS G ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.CODCCUSTO = G.CODCCUSTO AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = G.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN GCCUSTO AS L ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(C.CODCCUSTO,1,4) = L.CODCCUSTO AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = L.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN GCCUSTO AS M ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(C.CODCCUSTO,1,7) = M.CODCCUSTO AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = M.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN TTB3 AS H ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODTB3FAT = H.CODTB3FAT AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = H.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN TTB5 AS I ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.CODTB5FAT = I.CODTB5FAT " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN TTB5 AS J ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         J.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,4)+'.00.00' AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = J.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     LEFT JOIN TTB5 AS K ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         K.CODTB5FAT = SUBSTRING(I.CODTB5FAT,1,1)+'.00.00.00' AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = K.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     WHERE " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         /*PARAMETROS DECLARADOS E SETADOS NO INICIO*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA = " & vColigada & " AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODTMV IN(" & vMovs & ") AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.DATASAIDA BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' " & vbCrLf
        
'QUANDO INFORMA UMA FCE ESPECIFICA
        If vFCECC <> "" Then
                sqlCentroCusto = sqlCentroCusto & "                         AND(A.CODTB3FAT LIKE '" & vFCECC & "%' OR A.CAMPOLIVRE3 LIKE '" & vFCECC & "%')" & vbCrLf
        End If
'QUANDO INFORMA O CENTRO DE CUSTO OU PARTE DELE
        If vCustos <> "" Then
                sqlCentroCusto = sqlCentroCusto & "                         AND G.CODREDUZIDO LIKE '%" & vCustos & "%'" & vbCrLf
        End If
'QUANDO INFORMA UM PRODUTO ESPECIFICO
        If vProduto <> "" Then
                sqlCentroCusto = sqlCentroCusto & "                         AND E.NOMEFANTASIA LIKE '" & vProduto & "%'" & vbCrLf
        End If
        
        sqlCentroCusto = sqlCentroCusto & "                         /*PARAMETROS DECLARADOS E SETADOS NO INICIO*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     GROUP BY " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODCOLIGADA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         G.CODCCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         L.NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         M.NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         G.CODREDUZIDO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         G.NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.IDPRD, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.CODIGOPRD, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.NOMEFANTASIA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         E.CODUNDCONTROLE, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.QUANTIDADETOTAL, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         C.PRECOUNITARIO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         (C.QUANTIDADEARECEBER*C.PRECOUNITARIO), " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.DATASAIDA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODTMV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.IDMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.NUMEROMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CAMPOLIVRE3, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         A.CODTB3FAT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         H.DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(K.CODTB5FAT,1,1), " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         K.DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(J.CODTB5FAT,3,2), " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         J.DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         SUBSTRING(I.CODTB5FAT,6,2), " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                         I.DESCRICAO,A.RECCREATEDON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                     /*FIM DA PRIMEIRA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "                 ) AS CUSTO /*ALIAS DA QUERY DA PRIMEIRA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "             /*FIM DA SEGUNDA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         ) AS SEM_TRIBUTOS /*ALIAS DA QUERY DA SEGUNDA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "      " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     LEFT JOIN TITMMOV AS OUTRAS_DESP ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDMOV_ENTRADA = OUTRAS_DESP.IDMOV AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDPRD = OUTRAS_DESP.IDPRD AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         OUTRAS_DESP.CODCOLIGADA = SEM_TRIBUTOS.CODCOLIGADA " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     LEFT JOIN TTRBMOV AS TRIBUTOS ON " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         OUTRAS_DESP.IDMOV = TRIBUTOS.IDMOV AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         TRIBUTOS.CODTRB IN('IPI','ICMS','PIS001','COF001') AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         TRIBUTOS.CODCOLIGADA = SEM_TRIBUTOS.CODCOLIGADA AND " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         OUTRAS_DESP.NSEQITMMOV = TRIBUTOS.NSEQITMMOV " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     /*AQUI VOCE PODE FILTRAR PELO IDENTIFICADOR QUE VOCÊ QUIZER. TANTO DE ENTRADA QUANTO DE SAIDA*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     /* WHERE SEM_TRIBUTOS.IDMOV_ENTRADA = 26158 */ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     /*AQUI VOCE PODE FILTRAR PELO IDENTIFICADOR QUE VOCÊ QUIZER. TANTO DE ENTRADA QUANTO DE SAIDA*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     GROUP BY " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CODCOLIGADA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CODCCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CCUSTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.NOME, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDPRD, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.CODPRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.PRODUTO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.UND, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.QTDE, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.PRECOUNIT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.VALORTOTALITENS, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.PRECOUNITENT, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.ANO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.MES, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DIA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.TPMOV, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDMOV_SAIDA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.NUMNF, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.FCE_MANUAL, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.FCE_CC, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DESCRICAO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.COD_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DES_CUSTOGER, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.COD_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DES_SUBCENTRO1, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.COD_SUBCENTRO2, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.DES_SUBCENTRO, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.RECCREATEDON, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         TRIBUTOS.CODTRB, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         TRIBUTOS.ALIQUOTA, " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "         SEM_TRIBUTOS.IDMOV_ENTRADA" & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "     /*FIM DA TERCEIRA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & " ) AS DADOS /*ALIAS DA QUERY DA TERCEIRA PARTE*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "/*EXIBE OS IMPOSTOS E SEUS RESPECTIVOS PERCENTUAIS EM LINHA*/ " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "PIVOT (MAX(ALIQUOTA)  FOR COD_TRIBUTO IN ([IPI],[ICMS],[PIS001],[COF001])) AS COLUNAS_TRIBUTOS " & vbCrLf
        sqlCentroCusto = sqlCentroCusto & "ORDER BY ANO,MES,DIA,SUBCENTRO1,SUBCENTRO2,CCUSTO,COD_CUSTOGER,COD_SUBCENTRO1,COD_SUBCENTRO2"

'    End If
    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsCentroCusto.Open sqlCentroCusto, cnBanco, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set rsCentroCusto.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsCentroCusto
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (118)
    Screen.MousePointer = vbDefault
    
    rsCentroCusto.Close
    Set rsCentroCusto = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FCRCentroCusto.Hide
    Unload Me
    Set FCRCentroCusto = Nothing
    Form1.MousePointer = 0
    Form1.Label1.Visible = False
End Sub






