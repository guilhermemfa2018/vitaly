VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelatorios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios do sistema"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmRelatorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Parâmetros do Módulo Avaliador"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Frame Frame10 
         Caption         =   "Média geral"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   735
         Left            =   2880
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Label Label41 
            Caption         =   "Label41"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Formação escolar:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Cursos/treinamentos:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Habilidades:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Experiência:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mskCadMatriz 
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label40 
         Caption         =   "Label40"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label39 
         Caption         =   "Label39"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label38 
         Caption         =   "Label38"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   15
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRelatorios.frx":3469A
      PICN            =   "frmRelatorios.frx":346B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Tag             =   "Confirmar"
      ToolTipText     =   "Confirmar"
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRelatorios.frx":35390
      PICN            =   "frmRelatorios.frx":353AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCandidatos As New ADODB.Recordset
Private sqlCandidatos As String

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        chamaRel
    Case 2
        TiPo = False
        Unload Me
        Set frmRelatorios = Nothing
        Set frmRelatorios = Nothing
    End Select
End Sub

Private Sub Form_Load()
    listview_cabecalho
    MontaDadosLV
    criaTabTemp
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Lista de Relatorios", ListView1.Width / 1.1
    ListView1.View = lvwReport
End Sub

Private Sub MontaDadosLV()
    Set ItemLst = ListView1.ListItems.Add(, , "Grafico de competências")
    Set ItemLst = ListView1.ListItems.Add(, , "Programação de treinamentos anual")
    'Set ItemLst = ListView1.ListItems.Add(, , "Custo de treinamentos externos por peródo")
    Set ItemLst = ListView1.ListItems.Add(, , "Relação de cargos por treinamentos")
End Sub

Private Sub criaTabTemp()
On Error Resume Next
    'Criando uma tabela temporária global
    Dim rsTabTemp As New ADODB.Recordset
    Dim SqlTabTemp As String
    SqlTabTemp = "CREATE TABLE ##Tempglobal(id INT NOT NULL,CPF VARCHAR(50) NOT NULL,nomecolaborador VARCHAR(100) NOT NULL,departamento VARCHAR(100) NOT NULL, setor VARCHAR(100) NOT NULL, experiencia FLOAT NOT NULL, habilidade FLOAT NOT NULL, treinamento FLOAT NOT NULL, formacao FLOAT NOT NULL)"
    rsTabTemp.Open SqlTabTemp, cnBanco
End Sub

Private Function atualizaCandidatos()
    'FILTRA
    '1 = Colaborador
    '2 = Candidato
    atualizaCandidatos = True
    Dim rsDeletaTemp As New ADODB.Recordset
    Dim sqlDeletaTemp As String
    
    sqlDeletaTemp = "delete from ##Tempglobal"
    rsDeletaTemp.Open sqlDeletaTemp, cnBanco
    
    sqlCandidatos = "select a.id,a.cpf,a.nomecolaborador,d.nomedepartamento,e.nomesetor,c.codmatriz,f.nomecargo from tbcolaboradores as a inner join tbcolaboradoreshist as b " & _
    "on a.codcoligada = '" & vCodcoligada & "' and a.ativo = 'S' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join " & _
    "tbdepartamentos as d on c.coddepartamento = d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor " & _
    "inner join tbcargos as f on c.codcargo = f.codcargo order by a.id"
    rsCandidatos.Open sqlCandidatos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCandidatos.RecordCount = 0 Then
        rsCandidatos.Close
        Set rsCandidatos = Nothing
        atualizaCandidatos = False
        Exit Function
    End If
    
    If Not rsCandidatos.EOF Then
        While Not rsCandidatos.EOF '.Move(Val(Combo1.Text))
            txtCadMatriz(4) = rsCandidatos.Fields(5) ' Matriz
            Text1 = rsCandidatos.Fields(5) & rsCandidatos.Fields(6) ' Matrix+nome do cargo
            chkAvaliador(0).Value = 0
            chkAvaliador(1).Value = 0
            chkAvaliador(2).Value = 0
            chkAvaliador(3).Value = 0
            'For X = 0 To Len(rsCandidatos.Fields(5))
                chkAvaliador(0).Value = 1
                chkAvaliador(1).Value = 1
                chkAvaliador(2).Value = 1
                chkAvaliador(3).Value = 1
            'Next
            mskCadMatriz = rsCandidatos.Fields(1) ' CPF
            Avaliador "colaborador"
            GravaColaboradores
            rsCandidatos.MoveNext
        Wend
    End If
    rsCandidatos.Close
    Set rsCandidatos = Nothing
End Function

Private Sub GravaColaboradores()
On Error Resume Next
    Dim rsGravaColaboradores As New ADODB.Recordset
    Dim sqlGravaColaboradores As String
    Dim vIdent As Integer
    vIdent = rsCandidatos.Fields(0)
    
    sqlGravaColaboradores = "INSERT INTO ##Tempglobal(id,cpf,nomecolaborador,departamento,setor,experiencia,habilidade,treinamento,formacao) VALUES('" & rsCandidatos.Fields(0) & "','" & rsCandidatos.Fields(1) & "','" & rsCandidatos.Fields(2) & "','" & rsCandidatos.Fields(3) & "','" & rsCandidatos.Fields(4) & "','" & Replace(RemoveMask(Label37), ",", ".") & "','" & Replace(RemoveMask(Label38), ",", ".") & "','" & Replace(RemoveMask(Label39), ",", ".") & "','" & Replace(RemoveMask(Label41), ",", ".") & "')"
    rsGravaColaboradores.Open sqlGravaColaboradores, cnBanco
End Sub

Private Sub chamaRel()
On Error Resume Next
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If X = 1 Then
        If atualizaCandidatos = False Then
            MsgBox "Não há dados suficientes para gerar os gráficos", vbCritical, "SGCH"
            Exit Sub
        Else
            FCRGrafico.Show 1
        End If
    ElseIf X = 2 Then
        strAno = InputBox("Informe o ano", "SGCH")
        If StrPtr(strAno) = 0 Then
            MsgBox "Relatório Cancelado"
        Else
            If strAno <> "" Then
                FCRProgTrei.Show 1
            Else
                MsgBox "É necessário informar o ano"
            End If
        End If
    ElseIf X = 3 Then
            FCRTreinCargo.Show 1
    Else
        MsgBox "Rotina em desenvolvimento", vbCritical, "SGCH"
    End If
End Sub

