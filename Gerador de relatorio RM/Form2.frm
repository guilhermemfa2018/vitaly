VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form Form2 
   Caption         =   "Resumo de Produtos por FORNECEDOR"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport1

Private Sub Form_Load()
    Dim report1 As New CrystalReport1
    Dim rsRelAvTrei As New ADODB.Recordset
    Dim sqlRelAvTrei As String
    
    rsRelAvTrei.CursorLocation = adUseClient
    If vMovimento = "1.1.03" Then
        sqlRelAvTrei = "SELECT  '" & vNomeColigada & "'  AS NOMECOLIGADA,TMOV.NUMEROMOV AS NUMNF,FCFO.NOMEFANTASIA AS FORNECEDOR,TPRD.NOMEFANTASIA AS PRODUTO,TPRD.CODIGOPRD AS CODPRODUTO,TPRD.CODUNDCONTROLE AS UND,TITMMOV.QUANTIDADETOTAL AS QTDE,TITMMOV.QUANTIDADEARECEBER as QTDAREC,TITMMOV.QUANTIDADETOTAL - TITMMOV.QUANTIDADEARECEBER as QTDRECEBIDA,TITMMOV.PRECOUNITARIO AS PRECOUNIT,((TITMMOV.QUANTIDADETOTAL-TITMMOV.QUANTIDADEARECEBER)*TITMMOV.PRECOUNITARIO) AS TOTALITENS,TCPG.NOME AS CONDPAG,DATEPART(YEAR,TMOV.DATAEMISSAO)AS ANO,DATEPART(MONTH,TMOV.DATAEMISSAO)AS MES," & _
                     "DATEPART(DAY,TMOV.DATAEMISSAO)AS DIA,TMOV.CODTMV AS TPMOV,TMOV.CAMPOLIVRE3 AS FCE,TMOV.VALORBRUTO AS VLRPROD,TMOV.VALORLIQUIDO AS VLRLIQUIDO,TMOV.IDMOV,TMOV.status,TITMMOVCOMPL.LARGURA,TITMMOVCOMPL.COMPR FROM TMOV,FCFO,TITMMOV,TPRD,TCPG,TITMMOVCOMPL WHERE CODTMV IN('" & vMovimento & "')AND TMOV.CODCOLIGADA=FCFO.CODCOLIGADA AND TMOV.CODCFO=FCFO.CODCFO AND TMOV.CODCOLIGADA=TITMMOV.CODCOLIGADA AND TMOV.IDMOV=TITMMOV.IDMOV AND " & _
                     "TMOV.CODCPG=TCPG.CODCPG AND TMOV.SERIE='OC' AND TMOV.STATUS IN ('F','G','A') AND TITMMOV.IDPRD=TPRD.IDPRD AND TITMMOV.CODCOLIGADA=TPRD.CODCOLIGADA AND TITMMOV.IDMOV = TITMMOVCOMPL.IDMOV AND TITMMOV.NSEQITMMOV = TITMMOVCOMPL.NSEQITMMOV AND TMOV.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' AND FCFO.NOMEFANTASIA LIKE '" & "%" & vFornec & "%" & "' AND TMOV.CODCOLIGADA IN(" & vCodColigada & ") and FCFO.CODCOLIGADA IN(" & vCodColigada & ") and TITMMOV.CODCOLIGADA IN(" & vCodColigada & ") and tprd.CODCOLIGADA IN(" & vCodColigada & ") and  tcpg.CODCOLIGADA IN(" & vCodColigada & ") and  TITMMOVCOMPL.CODCOLIGADA IN(" & vCodColigada & ") ORDER BY ANO,MES,DIA"
    End If
    If vMovimento = "1.1.02" Then
        sqlRelAvTrei = "select  '" & vNomeColigada & "'  AS NOMECOLIGADA,a.NUMEROMOV AS NUMNF,b.NOMEFANTASIA AS FORNECEDOR,d.NOMEFANTASIA AS PRODUTO,d.CODIGOPRD AS CODPRODUTO,d.CODUNDCONTROLE AS UND,c.QUANTIDADETOTAL AS QTDE,c.QUANTIDADEARECEBER as QTDAREC,c.QUANTIDADETOTAL - c.QUANTIDADEARECEBER as QTDRECEBIDA,c.PRECOUNITARIO AS PRECOUNIT,((c.QUANTIDADETOTAL-c.QUANTIDADEARECEBER)*c.PRECOUNITARIO) AS TOTALITENS,e.NOME AS CONDPAG,DATEPART(YEAR,a.DATAEMISSAO)AS ANO," & _
                     "DATEPART(MONTH,a.DATAEMISSAO)AS MES,DATEPART(DAY,a.DATAEMISSAO)AS DIA,a.CODTMV AS TPMOV,a.CAMPOLIVRE3 AS FCE,a.VALORBRUTO AS VLRPROD,a.VALORLIQUIDO AS VLRLIQUIDO,a.IDMOV,a.status,f.largura,f.compr from TMOV as a left join FCFO as b on a.CODCOLIGADA=b.CODCOLIGADA AND a.CODCFO=b.CODCFO left join TITMMOV as c " & _
                     "on a.CODCOLIGADA=c.CODCOLIGADA AND a.IDMOV=c.IDMOV left join TPRD as d on c.IDPRD=d.IDPRD AND c.CODCOLIGADA=d.CODCOLIGADA left join TCPG AS e on a.CODCPG=e.CODCPG left join TITMMOVCOMPL as f on c.IDMOV = f.IDMOV and c.NSEQITMMOV = f.NSEQITMMOV where a.CODTMV IN('1.1.02') AND a.SERIE='SC' AND a.DATAEMISSAO BETWEEN '" & Format(vDataFilter1, "yyyy/mm/dd") & "' and '" & Format(vDataFilter2, "yyyy/mm/dd") & "' AND b.NOMEFANTASIA LIKE '" & "%" & vFornec & "%" & "' AND A.CODCOLIGADA IN(" & vCodColigada & ") and B.CODCOLIGADA IN(" & vCodColigada & ") and C.CODCOLIGADA IN(" & vCodColigada & ") and D.CODCOLIGADA IN(" & vCodColigada & ") and E.CODCOLIGADA IN(" & vCodColigada & ") and  F.CODCOLIGADA IN(" & vCodColigada & ") ORDER BY ANO,MES,DIA"
    End If
    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado

    rsRelAvTrei.Open sqlRelAvTrei, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsRelAvTrei.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsRelAvTrei
    Screen.MousePointer = vbHourglass

    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (102)
    Screen.MousePointer = vbDefault
    
    rsRelAvTrei.Close
    Set rsRelAvTrei = Nothing
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Hide
    Unload Me
    Set Form2 = Nothing
End Sub


