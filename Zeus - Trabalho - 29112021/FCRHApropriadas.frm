VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRHApropriadas 
   Caption         =   "ROP - Relatório Operacional da Produção"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRHApropriadas.frx":0000
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
Attribute VB_Name = "FCRHApropriadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err
    Dim report1 As New CRHApropriadas
    Dim rsHApropriadas As New ADODB.Recordset
    Dim sqlHApropriadas As String
    
    Dim crystalHApropriadas As New CRAXDRT.Application
    Dim ReportHApropriadas As CRAXDRT.Report
    
    'Linha abaixo - Altera texto do cabeçalho do relatório via código
    report1.Sections("ReportHeaderSection1").ReportObjects("Text17").SetText "ROP - RELATÓRIO OPERACIONAL DA PRODUÇÃO - PERÍODO: " & frmPrintRels.DTPicker1.Value & " - " & frmPrintRels.DTPicker2.Value & " - Semana: (" & frmPrintRels.Text2.Text & ")"
    
    rsHApropriadas.CursorLocation = adUseClient
    sqlHApropriadas = "select a.registro,a.nome,substring(a.centrocusto,6,10) as nomecc,CONVERT (VARCHAR, a.dataentrada, 103) as DataEntrada,CONVERT (VARCHAR, a.horaentrada, 108) as HoraEntrada,CONVERT (VARCHAR, a.datasaida, 103) as DataEntrada,CONVERT (VARCHAR, a.horasaida, 108) as HoraEntrada,a.idparada,a.nmparada,a.retrabalho, " & _
                      "case when a.TempoSRetrabalho = '0000:00' then '-' else a.TempoSRetrabalho end as TempoSRetrabalho,case when a.tempoCSRetrabalho = '0000:00' then '-' else a.tempoCSRetrabalho end as tempoCSRetrabalho,a.TempoTotalApropriacao,a.TempoTotalGeral,case when a.TempoPlanejadoCC is null then '-' else a.TempoPlanejadoCC end as TempoPlanejadoCC, " & _
                      "case when a.TempoPlanejadoTotal is null then '-' else a.TempoPlanejadoTotal end as TempoPlanejadoTotal,case when a.TempoTotalCarteira is null then '-' else a.TempoTotalCarteira end as TempoTotalCarteira,case when a.TempoGeralCarteira is null then '-' else a.TempoGeralCarteira end as TempoGeralCarteira,DATEPART(YEAR,a.dataentrada) AS ANO, " & _
                      "DATEPART(MONTH,a.dataentrada) AS MES,DATEPART(DAY,a.dataentrada) AS DIA,DATEPART(WK,a.dataentrada) as Semana,REPLACE(SUBSTRING(CONVERT (VARCHAR, a.tempoparada, 108),1,5),':',':') as TempoParada,REPLACE(a.tempototalcc,':',':'),case when a.tempototalpcc <> '0000:00' then REPLACE(a.tempototalpcc,':',':') else '-' end as tempototalpcc, " & _
                      "case when a.tempototalparada <> '0000:00' then REPLACE(a.tempototalparada,':',':') else '-' end as tempototalparada,case when a.tempototal <> '0000:00' then REPLACE(a.tempototal,':',':') else '-' end as tempototal,case when a.percentualtotalparada = '0' then '-' else RIGHT('0000000'+ CONVERT(VARCHAR,a.percentualtotalparada),7) end as tempototal, " & _
                      "case when a.ppsporcc = '00:00' then '-' else RIGHT('0000000'+ CONVERT(VARCHAR,a.ppsporcc),7) end as TempoPPS,case when a.AtrasoporCC = '00:00' then '-' else RIGHT('0000000'+ CONVERT(VARCHAR,a.AtrasoporCC),7) end as TempoAtraso,a.PPSTotal,a.AtrasoTotal,case when a.PPSeAtrasoPorCC <> '0000:00' then REPLACE(a.PPSeAtrasoPorCC,':',':') else '-' end as PPSeAtrasoPorCC, " & _
                      "a.PPSeAtrasoSoma,case when a.PPSRealPorCC = '0000:00' then '-' else a.PPSRealPorCC  end as PPSRealPorCC,a.PPSRealTotalPorCC,case when a.ExtraPPSRealPorCC = '00:00' then '-' else a.ExtraPPSRealPorCC  end as ExtraPPSRealPorCC,case when a.ExtraPPSRealTotalPorCC = '0000:00' then '-' else a.ExtraPPSRealTotalPorCC  end as ExtraPPSRealTotalPorCC," & _
                      "case when a.ExtraPPSRealSoma = '0000:00' then '-' else a.ExtraPPSRealSoma  end as ExtraPPSRealSoma,case when a.TempoTotalRealizado = '0000:00' then '-' else a.TempoTotalRealizado  end as TempoTotalRealizado from tbApropriaControle as a order by a.centrocusto,a.retrabalho,a.dataentrada"
    
    rsHApropriadas.Open sqlHApropriadas, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsHApropriadas.ActiveConnection = Nothing
    Set ReportHApropriadas = report1
    
    ReportHApropriadas.DiscardSavedData
    ReportHApropriadas.Database.SetDataSource rsHApropriadas
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (150)
    Screen.MousePointer = vbDefault
    rsHApropriadas.Close
    Set rsHApropriadas = Nothing
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

Private Sub Form_Unload(Cancel As Integer)
    FCRHApropriadas.Hide
    Unload Me
    Set FCRHApropriadas = Nothing
End Sub
