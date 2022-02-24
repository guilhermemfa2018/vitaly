VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRCRelatorio 
   Caption         =   "Ordem de Serviço"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "FCRCRelatorio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
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
      EnableGroupTree =   0   'False
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
Attribute VB_Name = "FCRCRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New CRrelatorio

Private Sub Form_Load()
On Error GoTo Err
    Dim report1 As New CRrelatorio
    Dim rsOS As New ADODB.Recordset
    Dim SqlOS As String
    
    Dim ReportOS As CRAXDRT.Report
    
    rsOS.CursorLocation = adUseClient

    SqlOS = SqlOS & "SELECT " & vbCrLf
    SqlOS = SqlOS & " A.NUMERO, " & vbCrLf
    SqlOS = SqlOS & " A.CLIENTE, " & vbCrLf
    SqlOS = SqlOS & " A.CONTATO, " & vbCrLf
    SqlOS = SqlOS & " A.TELEFONE, " & vbCrLf
    SqlOS = SqlOS & " A.EQUIPA, " & vbCrLf
    SqlOS = SqlOS & " A.MODELO, " & vbCrLf
    SqlOS = SqlOS & " A.MARCA, " & vbCrLf
    SqlOS = SqlOS & " A.ACESSORIOS, " & vbCrLf
    SqlOS = SqlOS & " A.SERIE, " & vbCrLf
    SqlOS = SqlOS & " A.SITUACAO, " & vbCrLf
    SqlOS = SqlOS & " A.DEFEITO1, " & vbCrLf
    SqlOS = SqlOS & " A.DEFEITO2, " & vbCrLf
    SqlOS = SqlOS & " A.SERVICO1, " & vbCrLf
    SqlOS = SqlOS & " A.SERVICO2, " & vbCrLf
    SqlOS = SqlOS & " A.SERVICO3, " & vbCrLf
    SqlOS = SqlOS & " A.SERVICO4, " & vbCrLf
    SqlOS = SqlOS & " A.SERVICO5, " & vbCrLf
    SqlOS = SqlOS & " A.OBSERVA, " & vbCrLf
    SqlOS = SqlOS & " A.TECNICO, " & vbCrLf
    SqlOS = SqlOS & " A.HORAENT, " & vbCrLf
    SqlOS = SqlOS & " A.DATAENT, " & vbCrLf
    SqlOS = SqlOS & " A.HORASAI, " & vbCrLf
    SqlOS = SqlOS & " A.DATASAI, " & vbCrLf
    SqlOS = SqlOS & " A.MAO_OBRA, " & vbCrLf
    SqlOS = SqlOS & " A.DESCONTO, " & vbCrLf
    SqlOS = SqlOS & " A.TOTAL, " & vbCrLf
    SqlOS = SqlOS & " A.VLRPROD, " & vbCrLf
    SqlOS = SqlOS & " A.PECAS, " & vbCrLf
    SqlOS = SqlOS & " B.EMPRESA, " & vbCrLf
    SqlOS = SqlOS & " B.RUA, " & vbCrLf
    SqlOS = SqlOS & " B.EMAIL, " & vbCrLf
    SqlOS = SqlOS & " B.NUM, " & vbCrLf
    SqlOS = SqlOS & " B.CID, " & vbCrLf
    SqlOS = SqlOS & " B.UF, " & vbCrLf
    SqlOS = SqlOS & " B.CEP, " & vbCrLf
    SqlOS = SqlOS & " B.BAI, " & vbCrLf
    SqlOS = SqlOS & " B.FAX AS TEL, " & vbCrLf
    SqlOS = SqlOS & " C.E_LOGO AS LOGO, " & vbCrLf
    SqlOS = SqlOS & " D.RUA AS CLI_RUA, " & vbCrLf
    SqlOS = SqlOS & " D.BAI AS CLI_BAI, " & vbCrLf
    SqlOS = SqlOS & " D.CID AS CLI_CID, " & vbCrLf
    SqlOS = SqlOS & " D.UF AS CLI_UF, " & vbCrLf
    SqlOS = SqlOS & " STR(D.CGC)  AS CLI_CNPJ, " & vbCrLf
    SqlOS = SqlOS & " D.CEP AS CLI_CEP " & vbCrLf
    SqlOS = SqlOS & " " & vbCrLf
    SqlOS = SqlOS & "FROM ORDEM AS A, DADOS AS B, LOGO AS C, CLIENTES AS D " & vbCrLf
    SqlOS = SqlOS & "WHERE"
    SqlOS = SqlOS & " A.NUMERO = " & vOS & " AND A.CODCLI = D.COD"
    
    rsOS.Open SqlOS, cnBancoDBF, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsOS.RecordCount = 0 Then
        rsOS.Close
        Set rsOS = Nothing
        MsgBox "Não existem dados a serem exibidos"
        Exit Sub
    End If
    Set rsOS.ActiveConnection = Nothing
    Set ReportOS = report1
    ReportOS.DiscardSavedData
    ReportOS.Database.SetDataSource rsOS
    Screen.MousePointer = vbHourglass
    'Report.RecordSelectionFormula = "{OrdemServico.os}= " & Val(varGlobal)
    CRViewer1.ReportSource = report1
    CRViewer1.ViewReport
    CRViewer1.Zoom (120)
    Screen.MousePointer = vbDefault
    rsOS.Close
    Set rsOS = Nothing
    Exit Sub
Err:
    MsgBox "Conexao não estabelecida com as tabelas de dados"
    Exit Sub
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub
