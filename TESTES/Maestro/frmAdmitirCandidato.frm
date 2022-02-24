VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdmitirCandidato 
   Caption         =   "Novo colaborador"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "frmAdmitirCandidato.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   30
      Top             =   5160
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
      MICON           =   "frmAdmitirCandidato.frx":0CCA
      PICN            =   "frmAdmitirCandidato.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   5160
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
      MICON           =   "frmAdmitirCandidato.frx":19C0
      PICN            =   "frmAdmitirCandidato.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame5 
      Caption         =   "Status"
      Height          =   1215
      Left            =   5400
      TabIndex        =   25
      Top             =   3840
      Width           =   2295
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Observação"
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   5175
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmAdmitirCandidato.frx":26B6
         Left            =   120
         List            =   "frmAdmitirCandidato.frx":26BD
         TabIndex        =   24
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Foto "
      Height          =   2415
      Index           =   0
      Left            =   5760
      TabIndex        =   17
      Top             =   120
      Width           =   1935
      Begin MSComDlg.CommonDialog cdlFoto 
         Left            =   1080
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   1635
         TabIndex        =   18
         Top             =   240
         Width           =   1695
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   2175
            Left            =   0
            Top             =   -120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   3836
            Image           =   "frmAdmitirCandidato.frx":26DC
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   1095
      Left            =   5400
      TabIndex        =   14
      Top             =   2640
      Width           =   2295
      Begin VB.TextBox txtNovoCol 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Tag             =   "Registro do novo colaborador"
         ToolTipText     =   "Registro do novo colaborador"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Registro nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações do candidato "
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Width           =   1335
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   60
            TabIndex        =   8
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Tag             =   "Matriz e cargo do colaborador"
         Text            =   "matriz - cargo"
         ToolTipText     =   "Matriz e cargo do colaborador"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label42 
         Caption         =   "Matriz/Cargo:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "CPF nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da requisição"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   5175
      Begin MAESTRO.chameleonButton cmdNovoCol 
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   28
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   2
         TX              =   "..."
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
         MICON           =   "frmAdmitirCandidato.frx":26F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtNovoCol 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   21
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNovoCol 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Matriz:"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "cargo:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Requisição nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "frmAdmitirCandidato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsNovoColaboradores As New ADODB.Recordset
Private SqlNovoColaboradores As String
Private rsReq As New ADODB.Recordset
Private sqlReq As String
Private rsLocal As New ADODB.Recordset

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        GravarDados
        gravaLog "CPF: " & txtNovoColaborador(0) & ", Registro: " & txtNovoCol(1), "Nome: " & txtNovoColaborador(1), "Média Geral: " & Label41 & ", Status: " & Label9
        carregaADP
        Unload Me
        Set frmAdmitirCandidato = Nothing
    Case 1
        ChamaGridReq
        If txtNovoCol(0) <> "" Then CarregaReq 1
        If txtNovoCol(2) <> "" Then CarregaReq 2
    Case 2
        Unload Me
        Set frmAdmitirCandidato = Nothing
    End Select
End Sub

Private Sub Form_Load()
    ResultPesq
End Sub

Private Sub ResultPesq()
    
    'SqlColaboradores = "Select * from tbColaboradores Where codcoligada = '" & vCodcoligada & "' and cpf = '" & Mid$(varGlobal, 1, 11) & "' and codcolaborador = '" & Mid$(varGlobal, 12, 10) & "' order by cpf"

    If apontaLV = 1 Then 'Candidato
        SqlNovoColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.id from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'candidato' and b.ativo = 'S' and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' and a.datarecisao is null order by a.cpf"
    ElseIf apontaLV = 0 Then 'Colaborador
        'SqlNovoColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.codcolaborador,a.homologacaonum,a.id from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'colaborador' and b.ativo = 'S' and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' and a.datarecisao is null order by a.cpf"
        SqlNovoColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.codcolaborador,a.homologacaonum,a.id from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' and a.codcolaborador = '" & Mid$(varGlobal, 12, 10) & "' inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'colaborador' and b.ativo = 'S' and a.datarecisao is null order by a.cpf"
    End If
    rsNovoColaboradores.Open SqlNovoColaboradores, cnBanco, adOpenKeyset, adLockReadOnly
    If rsNovoColaboradores.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir MeuLV.cmdconsulta(9).ToolTipText & " não encontrado", ok, critico, "Atenção"
    End If
    rsNovoColaboradores.Close
    Set rsNovoColaboradores = Nothing
End Sub

Private Sub CompoeControles()
On Error GoTo TrataErro1
    txtNovoColaborador(0).Text = Mid$(varGlobal, 1, 11)
    
    txtNovoColaborador(1).Text = rsNovoColaboradores.Fields(1)
    If apontaLV = 1 Then 'Candidato
        Label41 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
        txtNovoColaborador(2) = rsNovoColaboradores.Fields(5)
    ElseIf apontaLV = 0 Then 'Colaborador
        Label41 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(5)
        txtNovoCol(1).Enabled = False
        txtNovoCol(1) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(1)
        txtNovoColaborador(2) = rsNovoColaboradores.Fields(7)
    End If
    
    If Val(Label41) < MediaGlobal And Val(Label41) >= vAprovadoRest Then
        Label41.ForeColor = &H40C0&
        Label9.ForeColor = &H40C0&
        Label9.Caption = "Aprovado com restrição"
    ElseIf Val(Label41) < vAprovadoRest Then
        Label41.ForeColor = &HC0&
        Label9.ForeColor = &HC0&
        Label9.Caption = "Reprovado"
    ElseIf Val(Label41) >= MediaGlobal Then
        Label41.ForeColor = &H8000&
        Label9.ForeColor = &H8000&
        Frame4.Enabled = False
        Combo1.Enabled = False
        Label9.Caption = "Aprovado"
    End If
    If apontaLV = 0 Then 'Colaborador
        If Not IsNull(rsNovoColaboradores(6)) Then
            Label9.ForeColor = &HC0&
            Label9.Caption = "DEMITIDO"
        End If
    End If
    
    txtNovoColaborador(3) = Format(rsNovoColaboradores.Fields(2), "000000") & "-" & rsNovoColaboradores.Fields(3)
    Label53.Caption = rsNovoColaboradores.Fields(4)
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Sub ChamaGridReq()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select a.codrequisicao,b.codmatriz,d.nomecargo,b.numvagas,b.qtdocupada from tbrequisicoes as a inner join tbRequisicoesCargos as b on b.codrequisicao = a.codrequisicao inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.status = 'Aberto' and a.codcoligada = '" & vCodcoligada & "' order by a.codrequisicao"
    procnom = "nomecargo"
    procnom1 = "codcargo"
    campo = 0
    Campo1 = 1
    campo2 = 2
    Pesquisa = "Admissao"
    Load F
    F.Caption = "Pesquisa de Requisições"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        If Pesquisa <> 0 Then
            txtNovoCol(0).Text = Mid$(Pesquisa, 1, 6)
            txtNovoCol(2).Text = Mid$(Pesquisa, 7, 6)
        End If
        txtNovoCol(0).SetFocus
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub CarregaReq(campo As Integer)
    Dim X As Integer
    If campo = 1 Then sqlReq = "Select a.codrequisicao from tbrequisicoes as a inner join tbRequisicoesCargos as b on a.codcoligada = '" & vCodcoligada & "' and b.codrequisicao = a.codrequisicao inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.status = 'Aberto' and a.codrequisicao = '" & Val(txtNovoCol(0)) & "' order by a.codrequisicao"
    If campo = 2 Then sqlReq = "Select a.codrequisicao,b.codmatriz,d.nomecargo from tbrequisicoes as a inner join tbRequisicoesCargos as b on a.codcoligada = '" & vCodcoligada & "' and b.codrequisicao = a.codrequisicao inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.status = 'Aberto' and a.codrequisicao = '" & Val(txtNovoCol(0)) & "' and b.codmatriz= '" & Val(txtNovoCol(2)) & "' order by a.codrequisicao"
    rsReq.Open sqlReq, cnBanco, adOpenKeyset, adLockReadOnly
    If rsReq.RecordCount <= 0 Then
        If campo = 1 Then
            mobjMsg.Abrir "Requisição não encontrada", ok, critico, "Atenção"
            txtNovoCol(0).SetFocus
        End If
        If campo = 2 Then
            mobjMsg.Abrir "Cargo não encontrado para essa requisição", ok, critico, "Atenção"
            txtNovoCol(2).SetFocus
        End If
    Else
        If campo = 1 Then
            txtNovoCol(0).Text = Format(rsReq.Fields(0), "000000") & ""
            txtNovoCol(2).SetFocus
        End If
        If campo = 2 Then
            txtNovoCol(2).Text = Format(rsReq.Fields(1), "000000") & ""
            Label5 = rsReq.Fields(2)
            If txtNovoCol(2) <> Mid$(txtNovoColaborador(3), 1, 6) Then
                Label8.Caption = "O cargo selecionado é diferente do cargo pretendido pelo candidato"
                Label8.Visible = True
            Else
                Label8.Visible = False
            End If
        End If
    End If
    rsReq.Close
    Set rsReq = Nothing
End Sub

Private Sub txtNovoCol_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            CarregaReq 1
        End If
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            CarregaReq 2
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    fechaPDO
    If Label8.Visible = True Then
        mobjMsg.Abrir "A gravação não pode ser efetuada, devido a incompatibilidade das matrizes", ok, critico, "Atenção"
        Exit Sub
    End If
    Dim rsSalvarNovoCol As New ADODB.Recordset
    Dim SqlSalvarNovoCol As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim rsDadosTotvs As New ADODB.Recordset
    Dim SqlDBTotvs As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    'SqlSalvarNovoCol = "select * from tbcolaboradores where id = '" & txtNovoColaborador(2) & "'"
    'rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco, adOpenKeyset, adLockOptimistic
    
    
    SqlSalvarNovoCol = "Update tbColaboradores set codrequisicao = '" & Val(txtNovoCol(0)) & "',codcolaborador = '" & txtNovoCol(1) & "',ativo = 'S', obsadm = '" & Combo1 & "',tipo = 'colaborador' Where id= '" & Val(txtNovoColaborador(2)) & "' and codcoligada ='" & vCodcoligada & "'"
    rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
    
    'rsSalvarNovoCol.Fields(26) = Val(txtNovoCol(0)) 'codigo da requisicao
    'rsSalvarNovoCol.Fields(1) = txtNovoCol(1) 'Registro do novo colaborador
    'rsSalvarNovoCol.Fields(23) = "colaborador" 'Tipo
    'rsSalvarNovoCol.Fields(17) = "S" 'Ativo
    'If Combo1.Enabled = True Then rsSalvarNovoCol.Fields(28) = Combo1
    
    'rsSalvarNovoCol.Update
    
    SqlSalvar = "select * from tbcolaboradorescur where id = '" & txtNovoColaborador(2) & "' and codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsSalvar.EOF
        rsSalvar.Fields(1) = "colaborador"
        rsSalvar.MoveNext
    Wend
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    SqlSalvar = "select * from tbcolaboradoresesc where cpf = '" & txtNovoColaborador(0) & "' and codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsSalvar.EOF
        rsSalvar.Fields(1) = "colaborador"
        rsSalvar.MoveNext
    Wend
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    SqlSalvar = "select * from tbcolaboradoresexp where cpf = '" & txtNovoColaborador(0) & "' and codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsSalvar.EOF
        rsSalvar.Fields(1) = "colaborador"
        rsSalvar.MoveNext
    Wend
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    SqlSalvar = "select * from tbcolaboradoreshab where cpf = '" & txtNovoColaborador(0) & "' and codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsSalvar.EOF
        rsSalvar.Fields(1) = "colaborador"
        rsSalvar.MoveNext
    Wend
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    sqlDeletar = "Delete from tbColaboradoreshist where cpf = '" & txtNovoColaborador(0) & "' and ativo <> 'S' and tipo = 'candidato' and codcoligada ='" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "select * from tbcolaboradoreshist where cpf = '" & txtNovoColaborador(0) & "' and codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsSalvar.EOF
        rsSalvar.Fields(7) = "colaborador"
        rsSalvar.MoveNext
    Wend
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    'SALVAR CARGOS REQUISITADOR - LISTVIEW1
    
    If txtNovoCol(2) <> "" Then
        SqlSalvar = "Select * from tbRequisicoesCargos where codrequisicao = '" & Val(txtNovoCol(0)) & "' and codmatriz = '" & Val(txtNovoCol(2)) & "' and codcoligada ='" & vCodcoligada & "'"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
        Dim qtdVagaOcupadas As Integer
        qtdVagaOcupadas = rsSalvar.Fields(7) + 1
        rsSalvar.Fields(7) = qtdVagaOcupadas
        If qtdVagaOcupadas >= rsSalvar.Fields(2) Then
            rsSalvar.Fields(8) = "Fechado"
        End If
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    
    '>>>>>> GRAVAR CURSOS/TREINAMENTOS PENDENTES <<<<<<<<<
        GravaTreiPen
        'Se o parametro GeraIntr for "S" grava treinamentos introdutorios para os colaboradores
        If GeraIntr = "S" Then GravaTreiIntrodutorio
        If GeraObri = "S" Then GravaTreiObrigatorio
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'rsSalvarNovoCol.Close
    'Set rsSalvarNovoCol = Nothing
    
'************** TESTE DE INTEGRAÇÃO TOTVS
    If vIntegra = "S" Then
        If apontaLV = 0 Then
            SqlDBTotvs = "Select a.nomecolaborador,a.datanascimento,a.ctpsnumero,a.foto,b.sexo,b.grauinst,b.tipoadm,b.motadm,b.forreceb,b.situacao,b.tipofunc,b.hortrab,b.funcao,b.secao,b.contsind,b.rais,b.memsind " & _
            "from tbColaboradores as a LEFT join tbColaboradoresIntTotvs as b on a.codcoligada = '" & vCodcoligada & "' and a.id = b.id where a.id = '" & Val(txtNovoColaborador(2)) & "'"
            rsDadosTotvs.Open SqlDBTotvs, cnBanco, adOpenKeyset, adLockReadOnly
        ElseIf apontaLV = 1 Then
            SqlDBTotvs = "Select a.nomecolaborador,a.datanascimento,a.ctpsnumero,a.foto,b.sexo,b.grauinst,b.tipoadm,b.motadm,b.forreceb,b.situacao, " & _
            "b.tipofunc,b.hortrab,b.funcao,b.secao,b.contsind,b.rais,b.memsind from tbColaboradores as a left join tbColaboradoresIntTotvs as b on a.codcoligada = '" & vCodcoligada & "' and a.id = b.id where a.ativo = 'S' and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' and a.datarecisao is null order by a.cpf"
            rsDadosTotvs.Open SqlDBTotvs, cnBanco, adOpenKeyset, adLockReadOnly
        End If
        vDadosTotvs(0) = txtNovoCol(1) 'Chapa
        vDadosTotvs(1) = rsDadosTotvs.Fields(0) 'Nome do colaborador
        vDadosTotvs(2) = rsDadosTotvs.Fields(1) 'Data de nascimento
        vDadosTotvs(3) = rsDadosTotvs.Fields(2) 'Carteira de trabalho
        vDadosTotvs(4) = rsDadosTotvs.Fields(3) 'caminho foto
        vDadosTotvs(5) = rsDadosTotvs.Fields(4) 'sexo
        vDadosTotvs(6) = rsDadosTotvs.Fields(5) 'grau de instrução
        vDadosTotvs(7) = rsDadosTotvs.Fields(6) 'tipo de admissão
        vDadosTotvs(8) = rsDadosTotvs.Fields(7) 'motivo da admissão
        vDadosTotvs(9) = rsDadosTotvs.Fields(8) 'forma de recebimento
        vDadosTotvs(10) = rsDadosTotvs.Fields(9) 'Situação
        vDadosTotvs(11) = rsDadosTotvs.Fields(10) 'Tipo de funcionário
        vDadosTotvs(12) = rsDadosTotvs.Fields(11) 'horário de trabalho
        vDadosTotvs(13) = rsDadosTotvs.Fields(12) 'função
        vDadosTotvs(14) = rsDadosTotvs.Fields(13) 'seção
        vDadosTotvs(15) = rsDadosTotvs.Fields(14) 'contribuição sindical
        vDadosTotvs(16) = rsDadosTotvs.Fields(15) 'situação/rais
        vDadosTotvs(17) = rsDadosTotvs.Fields(16) 'membro sindicato
        For X = 0 To 17
            If vDadosTotvs(X) = "" Then
                mobjMsg.Abrir "Dados incompletos para exportar para o Corpore RM Labore. Favor conferir no cadastro", ok, critico, "Atenção"
                GoTo TrataErro
            End If
        Next
        GravaDadosDBTotvs txtNovoCol(1)
        
        rsDadosTotvs.Close
        Set rsDadosTotvs = Nothing
    End If
' **********************************
    cnBanco.CommitTrans
    mobjMsg.Abrir "Admissão realizada com sucesso", ok, informacao, "SGC"
    AtualizaListview
    Unload Me
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub fechaPDO()
    If vStatusPDO = "S" Then
        If Trim(vDecisao) <> "Aprovado" Then
            mobjMsg.Abrir "O PDO nº: " & Format(vNumPDO, "000000") & " NÃO FOI APROVADO ", ok, critico, "Atenção"
            'Remover Numero de PDO da tabela de colaboradores
            SqlPDOColab = "Update tbColaboradores set autorizacao = Null Where cpf = '" & Mid$(varGlobal, 1, 11) & "' and codcolaborador = '" & Mid$(varGlobal, 12, 10) & "' and codcoligada = '" & vCodcoligada & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
            Exit Sub
        Else
            'Remover Numero de PDO da tabela de colaboradores
            SqlPDOColab = "Update tbColaboradores set autorizacao = Null Where cpf = '" & Mid$(varGlobal, 1, 11) & "' and codcolaborador = '" & Mid$(varGlobal, 12, 10) & "' and codcoligada = '" & vCodcoligada & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
        End If
    End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If Label9 = "DEMITIDO" Then
        mobjMsg.Abrir "Não deve-se ATIVAR colaboradores com status de DEMISSÃO", ok, critico, "Atenção"
        Exit Function
    End If
    If txtNovoCol(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtNovoCol(1).Tag, ok, critico, "Atenção"
        Me.txtNovoCol(1).SetFocus
        Exit Function
    End If
    If Combo1.Enabled = True Then
        If Combo1.Text = "" Then
            mobjMsg.Abrir "Favor informar o motivo da admissão no campo de observação", ok, critico, "Atenção"
            Exit Function
        End If
    End If
    ValidaCampo = True
End Function

Private Sub GravaTreiPen()
    Dim rsGravaTreiPen As New ADODB.Recordset
    Dim SqlGravaTreiPen As String
    Dim rsPendentesCur As New ADODB.Recordset
    Dim SqlPendentesCur As String
    Dim contaID As Integer

    SqlGravaTreiPen = "Select a.codmatriz,a.codtreinamento,b.codtreinamento,b.cpf,a.codnivel from tbmatrizcur as a left join tbcolaboradorescur as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento  and b.codnivel >= a.codnivel and b.tipo = 'colaborador' and b.cpf = '" & txtNovoColaborador(0).Text & "' inner join tbcolaboradores as c " & _
    "on b.cpf = c.cpf and c.id = '" & Val(txtNovoColaborador(2).Text) & "' where a.codmatriz = '" & Val(Mid$(txtNovoColaborador(3), 1, 6)) & "' order by a.codtreinamento"
    rsGravaTreiPen.Open SqlGravaTreiPen, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlPendentesCur = "Select * from tbPendentesCur order by id"
    rsPendentesCur.Open SqlPendentesCur, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not rsPendentesCur.EOF Then
        rsPendentesCur.MoveLast
        contaID = rsPendentesCur.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsPendentesCur.Close
    Set rsPendentesCur = Nothing
    
    While Not rsGravaTreiPen.EOF
        'If Not IsNull(rsGravaTreiPen.Fields(2)) Then
        If IsNull(rsGravaTreiPen.Fields(2)) Then
            SqlPendentesCur = "Select * from tbPendentesCur where cpf = '" & txtNovoColaborador(0).Text & "' and codtreinamento= '" & rsGravaTreiPen.Fields(1) & "' and codcoligada = '" & vCodcoligada & "' order by id"
            rsPendentesCur.Open SqlPendentesCur, cnBanco, adOpenKeyset, adLockOptimistic
            If rsPendentesCur.RecordCount = 0 Then
                rsPendentesCur.AddNew
                rsPendentesCur.Fields(0) = txtNovoColaborador(0).Text
                rsPendentesCur.Fields(1) = rsGravaTreiPen.Fields(0)
                rsPendentesCur.Fields(2) = rsGravaTreiPen.Fields(1)
                rsPendentesCur.Fields(4) = "S"
                rsPendentesCur.Fields(5) = contaID
                rsPendentesCur.Fields(6) = "Pendente"
                rsPendentesCur.Fields(7) = 0
                If IsNull(rsGravaTreiPen.Fields(4)) Then rsPendentesCur.Fields(12) = 0 Else rsPendentesCur.Fields(12) = rsGravaTreiPen.Fields(4)
                rsPendentesCur.Fields(14) = vCodcoligada ' Codigo da coligada
                contaID = contaID + 1
            Else
                If rsPendentesCur.Fields(4) = "N" Then
                    rsPendentesCur.AddNew
                    rsPendentesCur.Fields(0) = txtNovoColaborador(0).Text
                    rsPendentesCur.Fields(1) = rsGravaTreiPen.Fields(0)
                    rsPendentesCur.Fields(2) = rsGravaTreiPen.Fields(1)
                    rsPendentesCur.Fields(4) = "S"
                    rsPendentesCur.Fields(5) = contaID
                    rsPendentesCur.Fields(6) = "Pendente"
                    rsPendentesCur.Fields(7) = 0
                    If IsNull(rsGravaTreiPen.Fields(4)) Then rsPendentesCur.Fields(8) = 0 Else rsPendentesCur.Fields(8) = rsGravaTreiPen.Fields(4)
                    rsPendentesCur.Fields(14) = vCodcoligada ' Codigo da coligada
                    contaID = contaID + 1
                Else
                End If
            End If
            rsPendentesCur.Update
            rsPendentesCur.Close
        End If
        rsGravaTreiPen.MoveNext
    Wend
    Set rsPendentesCur = Nothing
    
    rsGravaTreiPen.Close
    Set rsGravaTreiPen = Nothing
End Sub

Private Sub GravaTreiIntrodutorio()
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiInt As New ADODB.Recordset
    Dim SqlSelecionaTreiInt As String
    Dim rsGravaTreiInt As New ADODB.Recordset
    Dim SqlGravaTreiInt As String
    Dim contaID As Integer
    
    'LOCALIZAR SETOR DO COLABORADOR
    SqlAchaSetor = "select a.codsetor from tbsetores as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codsetor = b.codsetor where b.codmatriz = '" & Val(Mid$(txtNovoColaborador(3), 1, 6)) & "'"
    rsAchaSetor.Open SqlAchaSetor, cnBanco, adOpenKeyset, adLockReadOnly
        
    
'    If ListView5.ListItems.Count > 1 Then
'        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codsetor = '" & rsAchaSetor.Fields(0) & "'"
'    Else
'        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codsetor = 0 or codsetor = '" & rsAchaSetor.Fields(0) & "'"
'    End If
    SqlSelecionaTreiInt = "select * from tbTreinamentosint where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    rsSelecionaTreiInt.Open SqlSelecionaTreiInt, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao from tbPendentesCur where codcoligada ='" & vCodcoligada & "'"
    rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiInt.EOF Then
        rsGravaTreiInt.MoveLast
        contaID = rsGravaTreiInt.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiInt.Close
    Set rsGravaTreiInt = Nothing
    
    While Not rsSelecionaTreiInt.EOF
        SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel,codcoligada from tbPendentesCur where cpf = '" & txtNovoColaborador(0).Text & "' and codtreinamento ='" & rsSelecionaTreiInt.Fields(0) & "' and codcoligada ='" & vCodcoligada & "'"
        rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaTreiInt.RecordCount = 0 Then
            rsGravaTreiInt.AddNew
            rsGravaTreiInt.Fields(0) = txtNovoColaborador(0).Text
            rsGravaTreiInt.Fields(1) = Val(Mid$(txtNovoColaborador(3), 1, 6))
            rsGravaTreiInt.Fields(2) = rsSelecionaTreiInt.Fields(0)
            rsGravaTreiInt.Fields(4) = "S"
            rsGravaTreiInt.Fields(5) = contaID
            rsGravaTreiInt.Fields(6) = "Pendente"
            rsGravaTreiInt.Fields(7) = 0
            rsGravaTreiInt.Fields(8) = 0
            rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        Else
            If rsGravaTreiInt.Fields(4) = "N" Then
                rsGravaTreiInt.AddNew
                rsGravaTreiInt.Fields(0) = txtNovoColaborador(0).Text
                rsGravaTreiInt.Fields(1) = Val(Mid$(txtNovoColaborador(3), 1, 6))
                rsGravaTreiInt.Fields(2) = rsSelecionaTreiInt.Fields(0)
                rsGravaTreiInt.Fields(4) = "S"
                rsGravaTreiInt.Fields(5) = contaID
                rsGravaTreiInt.Fields(6) = "Pendente"
                rsGravaTreiInt.Fields(7) = 0
                rsGravaTreiInt.Fields(8) = 0
                rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            End If
        End If
        rsGravaTreiInt.Update
        rsGravaTreiInt.Close
        rsSelecionaTreiInt.MoveNext
    Wend
    Set rsGravaTreiInt = Nothing
    
    rsAchaSetor.Close
    Set rsAchaSetor = Nothing
    
    rsSelecionaTreiInt.Close
    Set rsSelecionaTreiInt = Nothing
End Sub

Private Sub GravaTreiObrigatorio()
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiObr As New ADODB.Recordset
    Dim SqlSelecionaTreiObr As String
    Dim rsGravaTreiObr As New ADODB.Recordset
    Dim SqlGravaTreiObr As String
    Dim contaID As Integer
    
    'LOCALIZAR SETOR DO COLABORADOR
    SqlAchaSetor = "select a.codsetor from tbsetores as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codsetor = b.codsetor where b.codmatriz = '" & Val(Mid$(txtNovoColaborador(3), 1, 6)) & "'"
    rsAchaSetor.Open SqlAchaSetor, cnBanco, adOpenKeyset, adLockReadOnly
        
    
'    If ListView5.ListItems.Count > 1 Then
'        SqlSelecionaTreiObr = "select * from tbTreinamentosObr where codsetor = '" & rsAchaSetor.Fields(0) & "'"
'    Else
'        SqlSelecionaTreiObr = "select * from tbTreinamentosObr where codsetor = 0 or codsetor = '" & rsAchaSetor.Fields(0) & "'"
'    End If
    SqlSelecionaTreiObr = "select * from tbTreinamentosObr where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    rsSelecionaTreiObr.Open SqlSelecionaTreiObr, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlGravaTreiObr = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao from tbPendentesCur where codcoligada = '" & vCodcoligada & "'"
    rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiObr.EOF Then
        rsGravaTreiObr.MoveLast
        contaID = rsGravaTreiObr.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiObr.Close
    Set rsGravaTreiObr = Nothing
    
    While Not rsSelecionaTreiObr.EOF
        SqlGravaTreiObr = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel,codcoligada from tbPendentesCur where cpf = '" & txtNovoColaborador(0).Text & "' and codtreinamento ='" & rsSelecionaTreiObr.Fields(0) & "' and codcoligada ='" & vCodcoligada & "'"
        rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaTreiObr.RecordCount = 0 Then
            rsGravaTreiObr.AddNew
            rsGravaTreiObr.Fields(0) = txtNovoColaborador(0).Text
            rsGravaTreiObr.Fields(1) = Val(Mid$(txtNovoColaborador(3), 1, 6))
            rsGravaTreiObr.Fields(2) = rsSelecionaTreiObr.Fields(0)
            rsGravaTreiObr.Fields(4) = "S"
            rsGravaTreiObr.Fields(5) = contaID
            rsGravaTreiObr.Fields(6) = "Pendente"
            rsGravaTreiObr.Fields(7) = 0
            rsGravaTreiObr.Fields(8) = 0
            rsGravaTreiObr.Fields(9) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        Else
            If rsGravaTreiObr.Fields(4) = "N" Then
                rsGravaTreiObr.AddNew
                rsGravaTreiObr.Fields(0) = txtNovoColaborador(0).Text
                rsGravaTreiObr.Fields(1) = Val(Mid$(txtNovoColaborador(3), 1, 6))
                rsGravaTreiObr.Fields(2) = rsSelecionaTreiObr.Fields(0)
                rsGravaTreiObr.Fields(4) = "S"
                rsGravaTreiObr.Fields(5) = contaID
                rsGravaTreiObr.Fields(6) = "Pendente"
                rsGravaTreiObr.Fields(7) = 0
                rsGravaTreiObr.Fields(8) = 0
                rsGravaTreiObr.Fields(9) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            End If
        End If
        rsGravaTreiObr.Update
        rsGravaTreiObr.Close
        rsSelecionaTreiObr.MoveNext
    Wend
    Set rsGravaTreiObr = Nothing
    
    rsAchaSetor.Close
    Set rsAchaSetor = Nothing
    
    rsSelecionaTreiObr.Close
    Set rsSelecionaTreiObr = Nothing
End Sub

Private Sub AtualizaListview()
'    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If apontaLV = 0 Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK"
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "OK"
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", ok, critico, "Atenção"
    Exit Sub
End Sub

