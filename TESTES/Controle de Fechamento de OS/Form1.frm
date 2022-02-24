VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Fechamento de Os's"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Configuração de conexão DB RM Sistemas "
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
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   8
         Text            =   "vigamax"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "sa"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Text            =   "zeus"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "SRV1002\CORPORERM"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Quantidade de OS's fechadas pelo sistema"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   6720
      Width           =   4335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   11040
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   6840
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colaboradores que não fecharam OS's"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   12255
      Begin MSComctlLib.ListView ListView1 
         Height          =   6015
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   10610
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vDataEnt As String, vChapa As String, vNome As String, vOs As String, vHSFunc As String, vNmCC As String, vOperacao As String
Private vCodHorario As Integer

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    FCRNaoFechou.Show 1
End Sub

Private Sub Form_Load()
    Conectar
    listview_cabecalho
    fechaOSsAbertas
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Dia", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Registro", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Os", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Contr. Sistema", ListView1.Width / 8
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub fechaOSsAbertas()
    Dim rsfechaOSsAbertas As New ADODB.Recordset
    Dim SqlfechaOSsAbertas As String
    Dim ItemLst As ListItem
    Dim vContador As Integer
    
    'Dim vHorario As String
    
    Dim rsHorarioFunc As New ADODB.Recordset
    Dim SqlHorarioFunc As String
    
    SqlfechaOSsAbertas = "select a.*,b.CODHORARIO,c.idos,c.nomecc,c.idoperacao from tbOsMov as a inner join CORPORERM.dbo.PFUNC as b on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA " & _
                         "LEFT join tbMPItens as c on a.codigobarra = c.codigobarra where datasai is null AND B.CODCOLIGADA >=5 order by A.CHAPA"
    rsfechaOSsAbertas.Open SqlfechaOSsAbertas, cnBanco, adOpenKeyset, adLockReadOnly
    If rsfechaOSsAbertas.RecordCount = 0 Then
        Exit Sub
    End If
    vContador = 0
    While Not rsfechaOSsAbertas.EOF
        vDataEnt = ""
        vChapa = ""
        vNome = ""
        vOs = ""
        vCodHorario = 0
        vHSFunc = ""
        vNmCC = ""
        vOperacao = ""
        vDataEnt = rsfechaOSsAbertas.Fields(3)
        vChapa = rsfechaOSsAbertas.Fields(1)
        If Not IsNull(rsfechaOSsAbertas.Fields(9)) Then vOs = rsfechaOSsAbertas.Fields(9)
        vCodHorario = rsfechaOSsAbertas.Fields(8)
        If Not IsNull(rsfechaOSsAbertas.Fields(10)) Then vNmCC = rsfechaOSsAbertas.Fields(10)
        If Not IsNull(rsfechaOSsAbertas.Fields(11)) Then vOperacao = rsfechaOSsAbertas.Fields(11)
        
        'SqlHorarioFunc = "DECLARE @Horario VARCHAR(4000) SET @Horario = '' " & _
        '                 "SELECT @Horario = RTRIM(@Horario) + RTRIM((REPLICATE('0', 2 - LEN(CAST((a.BATIDA /60) AS VARCHAR))) + CAST((a.BATIDA /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((a.BATIDA %60) AS VARCHAR))) + CAST((a.BATIDA %60) AS VARCHAR))) + ';' " & _
        '                 "FROM CORPORERM.dbo.ABATHOR as a where a.CODHORARIO = '" & Format(vCodHorario, "0000") & "' and a.INDICE = 1 AND A.BATIDA <> 0 GROUP BY A.CODHORARIO,A.INDICE, A.BATIDA " & _
        '                 "select a.CHAPA,b.NOME,c.CODHORARIO,c.INDICE,SUBSTRING(@Horario,1,5) ENT1,SUBSTRING(@Horario,7,5) SAI1,SUBSTRING(@Horario,13,5) ENT2,SUBSTRING(@Horario,19,5) SAI2 from CORPORERM.dbo.PFUNC as a " & _
        '                 "inner join CORPORERM.dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and a.CODPESSOA = b.CODIGO inner join CORPORERM.dbo.ABATHOR as c on a.CODHORARIO = c.CODHORARIO where a.CODHORARIO = '" & Format(vCodHorario, "0000") & "' and c.INDICE = 1 AND c.BATIDA <> 0 and a.CHAPA = '" & vChapa & "' " & _
        '                 "GROUP BY a.CHAPA,b.NOME,c.CODHORARIO,c.INDICE order by b.NOME"
        
        
        SqlHorarioFunc = "" & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "DECLARE @Horario VARCHAR(4000) " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "SET @Horario = '' " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "SELECT " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   @Horario = RTRIM(@Horario) + RTRIM((REPLICATE('0', 2 - LEN(CAST((a.BATIDA /60) AS VARCHAR))) + CAST((a.BATIDA /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((a.BATIDA %60) AS VARCHAR))) + CAST((a.BATIDA %60) AS VARCHAR))) + ';' " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "FROM CORPORERM.dbo.ABATHOR as a " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "where " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   a.CODHORARIO = '" & Format(vCodHorario, "0000") & "' and " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   a.INDICE = 1 AND " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   A.BATIDA <> 0 AND " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   a.CODCOLIGADA >= 5 " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "GROUP BY A.CODHORARIO,A.INDICE, A.BATIDA" & vbCrLf
        
        'SqlHorarioFunc = SqlHorarioFunc & "DECLARE @Horario VARCHAR(4000) SET @Horario = ''  SELECT " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   Horario = RTRIM(@Horario) + RTRIM((REPLICATE('0', 2 - LEN(CAST((a.BATIDA /60) AS VARCHAR))) + CAST((a.BATIDA /60) AS VARCHAR)+ ':' + REPLICATE('0', 2 - LEN(CAST((a.BATIDA %60) AS VARCHAR))) + CAST((a.BATIDA %60) AS VARCHAR))) + ';' " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "FROM CORPORERM.dbo.ABATHOR as a " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "where " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   a.CODHORARIO = '" & Format(vCodHorario, "0000") & "' and " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   a.INDICE = 1 AND " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   A.BATIDA <> 0 AND " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "   a.CODCOLIGADA >= 5 " & vbCrLf
        'SqlHorarioFunc = SqlHorarioFunc & "GROUP BY A.CODHORARIO,A.INDICE, A.BATIDA" & vbCrLf
        
        
        'rsHorarioFunc.Open SqlHorarioFunc, cnBanco, adLockReadOnly, adCmdText
        'If rsHorarioFunc.RecordCount > 0 Then
        '    vHorario = rsHorarioFunc.Fields(0)
        'End If
        
        
        SqlHorarioFunc = SqlHorarioFunc & "select " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CHAPA, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.NOME, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   c.CODHORARIO, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   c.INDICE, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   SUBSTRING('" & vHorario & "',1,5) ENT1, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   SUBSTRING('" & vHorario & "',7,5) SAI1, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   SUBSTRING('" & vHorario & "',13,5) ENT2, " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   SUBSTRING('" & vHorario & "',19,5) SAI2 " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "from CORPORERM.dbo.PFUNC as a " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "left join CORPORERM.dbo.PPESSOA as b on " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CODSITUACAO in('A','F','P','Z') and " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CODPESSOA = b.CODIGO " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "left join CORPORERM.dbo.ABATHOR as c on " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CODHORARIO = c.CODHORARIO AND " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CODCOLIGADA = c.CODCOLIGADA " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "where " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CODHORARIO = '" & Format(vCodHorario, "0000") & "' and " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   c.INDICE = 1 AND " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   c.BATIDA <> 0 and " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CHAPA = '" & vChapa & "' AND " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "   a.CODCOLIGADA >= 5 " & vbCrLf
        SqlHorarioFunc = SqlHorarioFunc & "GROUP BY a.CHAPA,a.NOME,c.CODHORARIO,c.INDICE order by A.CHAPA"
        
        rsHorarioFunc.Open SqlHorarioFunc, cnBanco, adLockReadOnly, adCmdText
        
        
        
        If rsHorarioFunc.RecordCount > 0 Then
            vNome = rsHorarioFunc.Fields(1)
            vHSFunc = rsHorarioFunc.Fields(7)
            'If Format(vDataEnt, "dd/mm/yyyy") = Format(Date - 1, "dd/mm/yyyy") Then
            Set ItemLst = ListView1.ListItems.Add(, , Format(vDataEnt, "dd/mm/yyyy"))
            ItemLst.SubItems(1) = "" & vChapa
            ItemLst.SubItems(2) = "" & vNome
            ItemLst.SubItems(3) = "" & Format(vOs, "000000000")
            ItemLst.SubItems(4) = "" & vHSFunc
            vContador = vContador + 1
            If Time > "23:30:00" Then
                'gravaParada
            End If
        End If
        rsHorarioFunc.Close
        rsfechaOSsAbertas.MoveNext
    Wend
    Label1 = vContador
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 1
    Me.ListView1.SortOrder = lvwAscending

    rsfechaOSsAbertas.Close
    Set rsfechaOSsAbertas = Nothing
    Set rsHorarioFunc = Nothing
End Sub

Private Sub gravaParada()
    Dim rsGravaOS As New ADODB.Recordset
    Dim SqlGravaOS As String
    
    Dim rsInsereOS As New ADODB.Recordset
    Dim SqlInsereOS As String
    
    SqlGravaOS = "Update tbOsMov set datasai = '" & Format(Date, "YYYY-MM-DD") & "',horasai = '" & vHSFunc & "', idparada ='9019' where chapa = '" & vChapa & "' and datasai is null"
    rsGravaOS.Open SqlGravaOS, cnBanco
    
    SqlInsereOS = "Insert into tbNaoFechaOS(dia,registro,nome,os,fechamento,nomecc,idoperacao) Values('" & Format(Date, "YYYY-MM-DD") & "','" & vChapa & "','" & vNome & "','" & vOs & "','" & vHSFunc & "','" & vNmCC & "','" & vOperacao & "')"
    rsInsereOS.Open SqlInsereOS, cnBanco
    
End Sub

