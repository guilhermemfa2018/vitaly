VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListaGRD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "GRD - Guia de Remessa de Documentos"
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Icon            =   "frmListaGRD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmListaGRD.frx":0CCA
   ScaleHeight     =   9660
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "vbNormal"
   Begin GRD.xVistaForm xVistaForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   688
      Caption         =   "GRD - Guia de Remessa de Documentos"
      DisplayIcon     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   16777215
      Style_Type      =   1
      EnableMaximiseButton=   0   'False
      FontItalic      =   0   'False
      FontSize        =   9,75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Icon            =   "frmListaGRD.frx":96C4
      ShowMaximiseButton=   0   'False
      ShowSytemTrayIcon=   -1  'True
      Style           =   1
      Transparency    =   -1  'True
   End
   Begin VB.Frame Frame7 
      Caption         =   "Configuração de conexão DB CORPORE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   12
         Text            =   "vigamax"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "sa"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   10
         Text            =   "CORPORE_ADM"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "LAPTOP-86MIAS4F"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   240
      Picture         =   "frmListaGRD.frx":A39E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir GRD"
      Top             =   8880
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Guia de Remessa de Documentos "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   6495
      Begin MSComDlg.CommonDialog cdg 
         Left            =   5760
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   11668
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   65280
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Relatórios GRD - Modelo: Arcelor Mittal "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Contrato nº"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   9120
      Width           =   5415
   End
End
Attribute VB_Name = "frmListaGRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CaminhoArquivo As String
Private NomeArquivo As String
Private pathArq As String
Private Plan As Object 'Aplicação Excel

Private Sub Command1_Click()
    ChamaGridContrato ("")
    chamaContrato Text1.Text
End Sub

Private Sub Command2_Click()
    AlteraListview
    SalvaXLS 1
End Sub

Private Sub Form_Load()
    xVistaForm1.Caption = "GRD - Guia de Remessa de Documentos" & " - Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    Conexao
    listview_cabecalho
    Text1.SetFocus
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 9
    ListView1.ColumnHeaders.Add , , "GRD", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Data Emissão", ListView1.Width / 5
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Function chamaContrato(vNome As String)
On Error GoTo Err
    Dim rschamaContrato As New ADODB.Recordset
    Dim SqlchamaContrato As String
    
    If vNome = "" Then
        SqlchamaContrato = "SELECT CODIGO,TITULO FROM ID_PRJ_CONTRATO order by codigo"
        rschamaContrato.Open SqlchamaContrato, cnBanco, adOpenKeyset, adLockReadOnly
    Else
        SqlchamaContrato = "SELECT CODIGO,TITULO FROM ID_PRJ_CONTRATO WHERE CODIGO like '%" & Text1.Text & "%' order by codigo"
        rschamaContrato.Open SqlchamaContrato, cnBanco, adOpenKeyset, adLockReadOnly
        If rschamaContrato.RecordCount > 1 Then
            rschamaContrato.Close
            Set rschamaContrato = Nothing
            ChamaGridContrato (SqlchamaContrato)
            SqlchamaContrato = "SELECT CODIGO,TITULO FROM ID_PRJ_CONTRATO WHERE CODIGO = '" & Pesquisa & "' order by codigo"
            rschamaContrato.Open SqlchamaContrato, cnBanco, adOpenKeyset, adLockReadOnly
        End If
        vNome = ""
    End If
    If Not rschamaContrato.EOF Then
        Text1.Text = rschamaContrato.Fields(0)
    Else
        MsgBox "Contrato não identificado no sistema", vbCritical, "Atenção"
        Text1.Text = rschamaContrato.Fields(0)
        Text1.SetFocus
    End If

    rschamaContrato.Close
    Set rschamaContrato = Nothing
    
    ListView1.ListItems.Clear
    chamaSQL "SELECT DGRD.IDGRD,IDDESCGRD,CONVERT(VARCHAR,DGRD.DATAEMISSAO,103) DATAEMISSAO FROM DGRD (NOLOCK) INNER JOIN DSPRJ (NOLOCK) ON DGRD.IDSPRJ = DSPRJ.IDSPRJ INNER JOIN DDOCGRD (NOLOCK) ON DGRD.IDGRD = DDOCGRD.IDGRD WHERE SUBSTRING(IDDESCGRD, 7, 4) = '" & Mid$(Text1.Text, 5, 4) & "' AND DGRD.DATAEMISSAO IS NOT NULL GROUP BY DGRD.IDGRD,IDDESCGRD,DGRD.DATAEMISSAO ORDER BY DGRD.IDDESCGRD desc"
    Compoe_Listview ListView1, Sqlp, ""

    
    Exit Function
Err:
    Exit Function
End Function

Private Sub ChamaGridContrato(vSqlp As String)
'On Error GoTo Err
    Dim F As New frmPesqger2
    If vSqlp = "" Then
        Sqlp = "SELECT CODIGO,TITULO FROM ID_PRJ_CONTRATO WHERE CODIGO  like '%" & Text1.Text & "%' order by codigo"
    Else
        Sqlp = vSqlp
        vSqlp = ""
    End If
    procnom = "titulo"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Contratos"
    Pesquisa = frmListaGRD.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "CODIGO=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If Pesquisa = "Pesquisa de Contratos" Then Pesquisa = ""
            Text1.Text = Pesquisa
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    If Err.Number = 3705 Then
        rsLocal.Close
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        Resume Next
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        If Text1.Text = "" Then
            If chamaContrato("") = False Then Exit Sub
        Else
            If chamaContrato(Text1.Text) = False Then Exit Sub 'txtEmprestimo(1).Text
        End If
    End If
End Sub

Private Function AlteraListview()
    AlteraListview = True
    Y = ListView1.ListItems.Count
    Dim vContaParaExclusao As Integer
    
    vContaParaExclusao = 0
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    
    If Y > 0 Then
        varGlobal = ListView1.ListItems.Item(X)
        NomeArquivo = ListView1.SelectedItem.ListSubItems.Item(1)
    End If
End Function

Private Sub SalvaXLS(Indice As Integer)
On Error GoTo testa_erro
    CaminhoArquivo = ""
    'NomeArquivo = ""
    CaminhoArquivo = pathArq 'Mid$(frmConfiguracao.txtCaminho, 1, Len(frmConfiguracao.txtCaminho) - Len("contratoNOVO.mdb"))
    'If Indice = 1 Then
        'NomeArquivo = NomeArquivo & ".xls"
    'ElseIf Indice = 2 Then
    '    NomeArquivo = "Fabricacao.xls"
    'End If
    
    cdg.Filter = "Planilha do Excel (*.xls)|*.xls"
    cdg.flags = cdlOFNHideReadOnly
    cdg.InitDir = CaminhoArquivo
    cdg.FileName = NomeArquivo
    pathArq = cdg.FileName
    cdg.ShowSave
    If Trim(pathArq) <> "" Then
        If Indice = 1 Then
            ExportaExcelEvolucao 'Plano de Carga
        'ElseIf Indice = 2 Then
        '    ExportaExcelEvolucao
        End If
    End If
    Exit Sub
testa_erro:
    If Err.Number = 32755 Then
        mobjMsg.Abrir "Procedimento cancelado", Ok, critico, "Atenção"
    End If
End Sub

Private Sub ExportaExcelEvolucao()
'On Error Resume Next
    'Dim vTCNC1 As String, vTCNC2 As String, vTGuil As String, vTTPuns As String, vTRosq As String, vTFRadial As String, vTFPrisma As String, vTFMag As String, vTSerraFita As String, vTCorte As String, vTDesemp As String, vTPrensa As String, vTMonC As String, vTMonN As String, vTSolC As String, vTSolN As String, vTAcabC As String, vTAcabN As String, vTCal As String, vTTrac As String
    
    Dim J As Integer, K As Integer, L As Integer, X As Integer
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")

    'CONTROI OS DADOS DA GRD SELECIONADA
    Dim rsGRD As New ADODB.Recordset
    Dim SqlGRD As String
    Dim vLin As Integer, vCol As Integer, vContaCol As Integer
    
    
    SqlGRD = "SELECT " & _
             "   IDDOCGRD,IDGRD,DDOCGRD.IDDOC,DDOCGRD.IDEXEC,DDOCGRD.IDDOCDEREVISAO,DDOCGRD.IDDESCDOC,DDOCGRD.DESCDOC,DDOCGRD.NCLI,DDOCGRD.NCLIFINAL, " & _
             "   DDOCGRD.IDTFRMT,DDOCGRD.QTDFLS,DDOCGRD.A1EQUIV,(SELECT A1EQUIV FROM DDOC WHERE DDOC.IDDOC = DDOCGRD.IDDOC) AS A1EQUIVDOC,DDOCGRD.REVDOC, " & _
             "   DDOCGRD.REVCLI,AMOTIVOEMIS.codmotivoemis,AMOTIVOEMIS.DESCRICAO AS MOTIVOEMIS,AMEIOEMIS.codmeioemis,AMEIOEMIS.DESCRICAO AS MEIOEMIS, " & _
             "   PERCENTFAT,PERCENTPROG,PERCENTREV,DDOCGRD.IDMED,DDOCGRD.VALORIMED,VALORTOTAL,IDCANCELADO,DDOCGRD.IDEAP,CONVERT(VARCHAR,DATAEMISSAO,103) AS DATAEMISSAO, " & _
             "   CONVERT(VARCHAR,DDOCGRD.DATACAD,103) AS DATACAD,DDOCGRD.IDUSERCAD,DATARETORNO,IDMOTRETORNO,OBSRETORNO,CONDPARTICIPACAO,A1PRODUZIDO, " & _
             "   PERCENTUALPARTICIPACAO,PREVISAOREVDOC,DEXECDOC.DESCEXEC,DIMED.IDMED,DIMED.IDDESCIMED,DIMED.DESCIMED,DDOCGRD.IDSETOREMITENTE,DSETOREMITENTE.NOME " & _
             "   SETOREMITENTE,DIMED.IDPRJ, CAST(DDOCGRD.REVCLI AS VARCHAR) + '.' +  CAST(DDOCGRD.REVDOC AS VARCHAR) as REVISAO,ddoc.NUFL " & _
             "FROM DDOCGRD (NOLOCK) " & _
             "INNER JOIN AMOTIVOEMIS (NOLOCK) ON AMOTIVOEMIS.CODMOTIVOEMIS = DDOCGRD.CODMOTIVOEMIS " & _
             "INNER JOIN AMEIOEMIS (NOLOCK) ON AMEIOEMIS.CODMEIOEMIS = DDOCGRD.CODMEIOEMIS " & _
             "LEFT JOIN DEXECDOC ON DDOCGRD.IDEXEC = DEXECDOC.IDEXEC " & _
             "LEFT JOIN DIMED ON DDOCGRD.IDMED = DIMED.IDMED " & _
             "LEFT JOIN DSETOREMITENTE ON DDOCGRD.IDSETOREMITENTE = DSETOREMITENTE.ID " & _
             "left join DDOC  (NOLOCK) ON DDOCGRD.IDDOC = DDOC.IDDOC " & _
             "WHERE IDGRD = " & varGlobal
    rsGRD.Open SqlGRD, cnBanco, adOpenKeyset, adLockReadOnly


    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\Modelo de GRD - Arcelor Mittal.xls"
    Plan.Visible = True
    Plan.UserControl = False

    
    'PREENCHE CÉLULAS DESEJADAS
    '**********************************************************************
    With Plan
        .Range("GRD!J" & 3).Value = rsGRD.Fields(27) ' Data de envio da GRD
        .Range("GRD!M" & 3).Value = NomeArquivo  ' Numero da GRD
    End With
    
    Dim antes As String, depois As String, Arquivo As String
    
    J = 22
    While Not rsGRD.EOF
        Arquivo = rsGRD.Fields(40)
        antes = Mid(Arquivo, 1, InStr(Arquivo, "/") - 1)
        depois = Mid(Arquivo, InStr(Arquivo, "/") + 1, Len(Arquivo))
        
        With Plan
            .Range("C" & J).Value = rsGRD.Fields(7) ' Numero do cliente
            .Range("D" & J).Value = rsGRD.Fields(9) ' formato
            .Range("E" & J).Value = rsGRD.Fields(13) ' Revisão do documento
            .Range("F" & J).Value = rsGRD.Fields(45) ' Quantidade de folhas
            .Range("G" & J).Value = rsGRD.Fields(39) ' codigo do item
            .Range("H" & J).Value = depois ' descricao
            .Range("K" & J).Value = antes ' tipo
            .Range("L" & J).Value = rsGRD.Fields(11) ' A1 Equivalente
        End With
        J = J + 1
        rsGRD.MoveNext
    Wend
    rsGRD.Close




    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.ActiveWorkbook.SaveAs cdg.FileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    
    'Plan.Close = True
    Set Plan = Nothing
    'Plan.Quit
    MsgBox "Dados exportados com sucesso", vbInformation, "Atenção"
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, O MSOffice não esta instalado nesse computador!", vbCritical, "Atenção"
    Exit Sub
End Sub

Private Sub xVistaForm1_Execute(ByVal ID As Long)
    Me.Text1.SetFocus
End Sub
