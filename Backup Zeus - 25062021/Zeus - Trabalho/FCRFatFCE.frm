VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRFatFCE 
   Caption         =   "Faturamento por FCE - Analítico"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FCRFatFCE.frx":0000
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
Attribute VB_Name = "FCRFatFCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRFatFCE

Private Sub Form_Load()
    'On Error Resume Next
    Dim report1 As New CRFatFCE
    Dim rsFatFCE As New ADODB.Recordset
    Dim sqlFatFCE As String
    Dim SomaPesoVendas As Double
    Dim SomaValorVendas As Double
    Dim AdiantamentoValor As Double
    Dim AdiantamentoaReceber As Double
    Dim ValoraFaturar As Double
    Dim PesoaReceber As TextObject
    
    Dim rsTbTemp As New ADODB.Recordset
    Dim sqlTbTemp As String
    
    rsFatFCE.CursorLocation = adUseClient
    
    strAno = Mid$(varGlobal, 1, 4)
    'report1.Sections("PageHeaderSection1").ReportObjects("Text1").SetText "RELATÓRIO DE FATURAMENTO - FCE: " '& strAno
    
    If vQualquerDado(1, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text40").SetText vQualquerDado(1, 1)
    If vQualquerDado(1, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text41").SetText vQualquerDado(1, 2)
    If vQualquerDado(1, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text42").SetText Format(vQualquerDado(1, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(1, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text43").SetText "R$ " & Format(vQualquerDado(1, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(1, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text28").SetText vQualquerDado(1, 5)
    If vQualquerDado(1, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text78").SetText vQualquerDado(1, 6) & "% - " & vQualquerDado(1, 7)
    
    If vQualquerDado(2, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text44").SetText vQualquerDado(2, 1)
    If vQualquerDado(2, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text45").SetText vQualquerDado(2, 2)
    If vQualquerDado(2, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text46").SetText Format(vQualquerDado(2, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(2, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text47").SetText "R$ " & Format(vQualquerDado(2, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(2, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text29").SetText vQualquerDado(2, 5)
    If vQualquerDado(2, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text79").SetText vQualquerDado(2, 6) & "% - " & vQualquerDado(2, 7)
    
    If vQualquerDado(3, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text48").SetText vQualquerDado(3, 1)
    If vQualquerDado(3, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text49").SetText vQualquerDado(3, 2)
    If vQualquerDado(3, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text50").SetText Format(vQualquerDado(3, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(3, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text51").SetText "R$ " & Format(vQualquerDado(3, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(3, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text60").SetText vQualquerDado(3, 5)
    If vQualquerDado(3, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text80").SetText vQualquerDado(3, 6) & "% - " & vQualquerDado(3, 7)
    
    If vQualquerDado(4, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text52").SetText vQualquerDado(4, 1)
    If vQualquerDado(4, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text53").SetText vQualquerDado(4, 2)
    If vQualquerDado(4, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text54").SetText Format(vQualquerDado(4, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(4, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text55").SetText "R$ " & Format(vQualquerDado(4, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(4, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text73").SetText vQualquerDado(4, 5)
    If vQualquerDado(4, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text81").SetText vQualquerDado(4, 6) & "% - " & vQualquerDado(4, 7)
    
    If vQualquerDado(5, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text56").SetText vQualquerDado(5, 1)
    If vQualquerDado(5, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text57").SetText vQualquerDado(5, 2)
    If vQualquerDado(5, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text58").SetText Format(vQualquerDado(5, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(5, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text59").SetText "R$ " & Format(vQualquerDado(5, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(5, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text74").SetText vQualquerDado(5, 5)
    If vQualquerDado(5, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text82").SetText vQualquerDado(5, 6) & "% - " & vQualquerDado(5, 7)
    
    
'-------------------------------
    
    If vQualquerDado(6, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text99").SetText vQualquerDado(6, 1)
    If vQualquerDado(6, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text100").SetText vQualquerDado(6, 2)
    If vQualquerDado(6, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text101").SetText Format(vQualquerDado(6, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(6, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text102").SetText "R$ " & Format(vQualquerDado(6, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(6, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text103").SetText vQualquerDado(6, 5)
    If vQualquerDado(6, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text104").SetText vQualquerDado(6, 6) & "% - " & vQualquerDado(6, 7)
    
    If vQualquerDado(7, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text105").SetText vQualquerDado(7, 1)
    If vQualquerDado(7, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text106").SetText vQualquerDado(7, 2)
    If vQualquerDado(7, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text107").SetText Format(vQualquerDado(7, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(7, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text108").SetText "R$ " & Format(vQualquerDado(7, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(7, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text109").SetText vQualquerDado(7, 5)
    If vQualquerDado(7, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text110").SetText vQualquerDado(7, 6) & "% - " & vQualquerDado(7, 7)
    
    If vQualquerDado(8, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text111").SetText vQualquerDado(8, 1)
    If vQualquerDado(8, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text113").SetText vQualquerDado(8, 2)
    If vQualquerDado(8, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text114").SetText Format(vQualquerDado(8, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(8, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text115").SetText "R$ " & Format(vQualquerDado(8, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(8, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text116").SetText vQualquerDado(8, 5)
    If vQualquerDado(8, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text117").SetText vQualquerDado(8, 6) & "% - " & vQualquerDado(8, 7)
    
    If vQualquerDado(9, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text118").SetText vQualquerDado(9, 1)
    If vQualquerDado(9, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text119").SetText vQualquerDado(9, 2)
    If vQualquerDado(9, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text120").SetText Format(vQualquerDado(9, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(9, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text121").SetText "R$ " & Format(vQualquerDado(9, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(9, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text122").SetText vQualquerDado(9, 5)
    If vQualquerDado(9, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text123").SetText vQualquerDado(9, 6) & "% - " & vQualquerDado(9, 7)
    
    If vQualquerDado(10, 1) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text124").SetText vQualquerDado(10, 1)
    If vQualquerDado(10, 2) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text125").SetText vQualquerDado(10, 2)
    If vQualquerDado(10, 3) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text126").SetText Format(vQualquerDado(10, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(10, 4) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text127").SetText "R$ " & Format(vQualquerDado(10, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(10, 5) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text128").SetText vQualquerDado(10, 5)
    If vQualquerDado(10, 6) <> "" Then report1.Sections("ReportFooterSection1").ReportObjects("Text129").SetText vQualquerDado(10, 6) & "% - " & vQualquerDado(10, 7)
    
'-------------------------------
    
    
    If vQualquerDado(1, 6) <> "" Then
        AdiantamentoValor = Format(vQualquerDado(1, 4) * vQualquerDado(1, 6) / 100, "#,##0.00;(#,##0.00)")
        If vQualquerDado(2, 6) <> "" Then
            AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(2, 4) * vQualquerDado(2, 6) / 100, "#,##0.00;(#,##0.00)")
            If vQualquerDado(3, 6) <> "" Then
                AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(3, 4) * vQualquerDado(3, 6) / 100, "#,##0.00;(#,##0.00)")
                If vQualquerDado(4, 6) <> "" Then
                    AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(4, 4) * vQualquerDado(4, 6) / 100, "#,##0.00;(#,##0.00)")
                    If vQualquerDado(5, 6) <> "" Then
                        AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(5, 4) * vQualquerDado(5, 6) / 100, "#,##0.00;(#,##0.00)")
                        If vQualquerDado(6, 6) <> "" Then
                            AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(6, 4) * vQualquerDado(6, 6) / 100, "#,##0.00;(#,##0.00)")
                            If vQualquerDado(7, 6) <> "" Then
                                AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(7, 4) * vQualquerDado(7, 6) / 100, "#,##0.00;(#,##0.00)")
                                If vQualquerDado(8, 6) <> "" Then
                                    AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(8, 4) * vQualquerDado(8, 6) / 100, "#,##0.00;(#,##0.00)")
                                    If vQualquerDado(9, 6) <> "" Then
                                        AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(9, 4) * vQualquerDado(9, 6) / 100, "#,##0.00;(#,##0.00)")
                                        If vQualquerDado(10, 6) <> "" Then
                                            AdiantamentoValor = AdiantamentoValor + Format(vQualquerDado(10, 4) * vQualquerDado(10, 6) / 100, "#,##0.00;(#,##0.00)")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    report1.Sections("ReportFooterSection1").ReportObjects("Text83").SetText "R$ " & Format(AdiantamentoValor, "#,##0.00;(#,##0.00)")
'    report1.Sections("ReportFooterSection1").ReportObjects("Text95").SetText "R$ " & Format(AdiantamentoValor, "#,##0.00;(#,##0.00)")
    
    If vQualquerDado(1, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(1, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(2, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(2, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(3, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(3, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(4, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(4, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(5, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(5, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(6, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(6, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(7, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(7, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(8, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(8, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(9, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(9, 3), "#,##0.00;(#,##0.00)")
    If vQualquerDado(10, 3) <> "" Then SomaPesoVendas = SomaPesoVendas + Format(vQualquerDado(10, 3), "#,##0.00;(#,##0.00)")
    
    
    If vQualquerDado(1, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(1, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(2, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(2, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(3, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(3, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(4, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(4, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(5, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(5, 4), "#,##0.00;(#,##0.00)")
    
    If vQualquerDado(6, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(6, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(7, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(7, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(8, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(8, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(9, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(9, 4), "#,##0.00;(#,##0.00)")
    If vQualquerDado(10, 4) <> "" Then SomaValorVendas = SomaValorVendas + Format(vQualquerDado(10, 4), "#,##0.00;(#,##0.00)")
    
    
    
    'Valor do Adiantamento
    If vQualquerDado(20, 1) <> "" Then
'        report1.Sections("ReportFooterSection1").ReportObjects("Text98").SetText "R$ " & Format(vQualquerDado(20, 1), "#,##0.00;(#,##0.00)")
        AdiantamentoaReceber = AdiantamentoValor - Format(vQualquerDado(20, 1), "#,##0.00;(#,##0.00)")
'        report1.Sections("ReportFooterSection1").ReportObjects("Text96").SetText Format(AdiantamentoValor - vQualquerDado(20, 1), "#,##0.00;(#,##0.00)")
    Else
        'report1.Sections("ReportFooterSection1").ReportObjects("Text96").SetText "R$ " & Format(AdiantamentoValor, "#,##0.00;(#,##0.00)")
    End If
    
    
    report1.Sections("ReportFooterSection1").ReportObjects("Text75").SetText Format(SomaPesoVendas, "#,##0.00;(#,##0.00)")
    report1.Sections("ReportFooterSection1").ReportObjects("Text27").SetText "R$ " & Format(SomaValorVendas, "#,##0.00;(#,##0.00)")
    report1.Sections("ReportFooterSection1").ReportObjects("Text94").SetText "R$ " & Format(SomaValorVendas, "#,##0.00;(#,##0.00)")
    report1.Sections("ReportFooterSection1").ReportObjects("Text132").SetText Format(SomaPesoVendas, "#,##0.00;(#,##0.00)")

    
'    report1.Sections("ReportFooterSection1").ReportObjects("Text97").SetText "R$ " & Format(SomaValorVendas - vQualquerDado(20, 2), "#,##0.00;(#,##0.00)")
'    report1.Sections("ReportFooterSection1").ReportObjects("Text134").SetText Format(SomaPesoVendas - vQualquerDado(20, 3), "#,##0.00;(#,##0.00)")

    sqlFatFCE = "select A.CODTMV,A.IDMOV,E.DESCRICAO,A.NUMEROMOV,C.ROMAVIGA,A.DATAEMISSAO,A.QUANTIDADE,A.PESOLIQUIDO,A.PESOBRUTO,A.VALORLIQUIDO,A.VALORBRUTO,A.CODUSUARIO AS USU_CRIACAO,B.NOME AS COND_PAG," & _
                "    CASE WHEN A.STATUS = 'P' THEN 'Parcialmente Quitado' " & _
                "         WHEN A.STATUS = 'C' THEN 'Cancelado'" & _
                "         WHEN A.STATUS = 'A' THEN 'Pendente/Faturar'" & _
                "         WHEN A.STATUS = 'Q' THEN 'Quitado'" & _
                "         WHEN A.STATUS = 'F' THEN 'Receber/A pagar'" & _
                "         ELSE 'NÃO IDENTIFICADO'" & _
                "    END AS STATUS" & _
                ",D.DATABAIXA,D.DATAVENCIMENTO,D.DATAPAG,isnull(D.VALORBAIXADO,0)+isnull(D.VALORADIANTAMENTO,0) as VALORBAIXADO,D.VALORORIGINAL,D.USUARIO AS USU_FINANC,D.NUMERODOCUMENTO,D.HISTORICO,isnull(D.VALORADIANTAMENTO,0) as VALORADIANTAMENTO,(isnull(D.VALORBAIXADO,0)+isnull(D.VALORADIANTAMENTO,0))-isnull(D.VALORADIANTAMENTO,0) as VALORFATURADO from " & vBancoTotvs & ".dbo.TMOV as a inner join " & vBancoTotvs & ".dbo.TCPG as  B ON A.CODCPG = B.CODCPG " & _
                "LEFT JOIN " & vBancoTotvs & ".dbo.TMOVCOMPL AS C ON A.IDMOV = C.IDMOV LEFT JOIN " & vBancoTotvs & ".dbo.FLAN AS D ON A.IDMOV = D.IDMOV INNER JOIN " & vBancoTotvs & ".dbo.TTB3 as E on A.CODTB3FAT = E.CODTB3FAT " & _
                "where a.CODTB3FAT = '" & strAno & "' and a.CODTMV in ('2.2.01','2.2.05','2.2.25','1.2.15','1.2.17') order by a.CODTMV desc,a.IDMOV,a.NUMEROMOV,d.NUMERODOCUMENTO"

    cnBanco.CommandTimeout = 0 'Tempo de espera do banco indeterminado
    rsFatFCE.Open sqlFatFCE, cnBanco, adOpenKeyset, adLockReadOnly, adCmdText
    
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
    Set FCRFatFCE = Nothing
    Unload Me
    'Form1.MousePointer = 0
    'Form1.Label1.Visible = False
End Sub
