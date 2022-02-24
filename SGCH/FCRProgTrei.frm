VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form FCRProgTrei 
   Caption         =   "Programação anual de cursos/treinamentos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FCRProgTrei.frx":0000
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
Attribute VB_Name = "FCRProgTrei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CRProgTrei
Dim sqlProgTrei As String

Private Sub Form_Load()
    Dim report1 As New CRProgTrei
    Dim rsProgTrei As New ADODB.Recordset
    
    rsProgTrei.CursorLocation = adUseClient
    sqlProgTrei = "SET LANGUAGE 'Portuguese'"
    rsProgTrei.Open sqlProgTrei, cnBanco
    sqlProgTrei = "select d.nometreinamento,e.nomeinstrutor,c.nomecolaborador,YEAR(b.datainicio) Ano,MONTH(b.datainicio) NumMes,Substring(DATENAME(MONTH,b.datainicio),1,3) as NomeMes," & _
             "DAY(b.datainicio) as Dia,b.datainicio,b.datafim,b.status,d.origem,F.logo,d.valor from tbPendentesCur as a inner join tbprogramacao as b on a.codprogramacao = b.codprogramacao inner join tbcolaboradores as c " & _
             "on a.cpf = c.cpf inner join tbtreinamentos as d on a.codtreinamento = d.codtreinamento inner join tbProgramacaoInstrutores as e on b.codprogramacao = e.codprogramacao inner join tbDadosEmpresa as f on f.codcoligada = '" & vCodColigada & "' order by Ano,NumMes,Dia,d.nometreinamento"
    rsProgTrei.Open sqlProgTrei, cnBanco, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsProgTrei.ActiveConnection = Nothing
    Set Report = report1
    Report.DiscardSavedData
    Report.Database.SetDataSource rsProgTrei
    Screen.MousePointer = vbHourglass
    Report.RecordSelectionFormula = "{ProgTrei.ano}= " & Val(strAno)
    
    If Not rsProgTrei.EOF Then rsProgTrei.MoveFirst
    rsProgTrei.Find "ano=" & "'" & strAno & "'"
    If rsProgTrei.EOF Then
        MsgBox "Não Exitem Registros para o ano de: " & strAno
        rsProgTrei.Close
        Set rsProgTrei = Nothing
        Set Report = Nothing
        Set report1 = Nothing
        Unload Me
    Else
        CRViewer1.ReportSource = report1
        CRViewer1.ViewReport
        CRViewer1.Zoom (88)
        Screen.MousePointer = vbDefault
        rsProgTrei.Close
        Set rsProgTrei = Nothing
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'CRProgTrei.Hide
    'FCRProgTrei.Hide
    Unload Me
    Set CRProgTrei = Nothing
End Sub

