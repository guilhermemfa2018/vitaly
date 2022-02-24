VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRetorna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retorno de afastamento"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "frmRetorna.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Dados do colaborador "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtRetorno 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRetorna.frx":0CCA
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRetorno 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtRetorno 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox txtRetorno 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   7815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dias afastado "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         TabIndex        =   19
         Top             =   240
         Width           =   2415
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetorna.frx":0D30
            TabIndex        =   20
            Top             =   360
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   222756865
         CurrentDate     =   41318
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRetorna.frx":0D90
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRetorna.frx":0E10
         TabIndex        =   22
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmRetorna.frx":0E88
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRetorna.frx":0EFE
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   222756865
         CurrentDate     =   41318
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmRetorna.frx":0F6E
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Treinamentos propostos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   8175
      Begin VB.ComboBox cboRetorno 
         Height          =   315
         Index           =   5
         Left            =   6000
         TabIndex        =   9
         Top             =   450
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   6000
         OleObjectBlob   =   "frmRetorna.frx":0FEC
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtRetorno 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   7
         Tag             =   "Nome do responsável pelo setor"
         ToolTipText     =   "Nome do responsável pelo setor"
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtRetorno 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Tag             =   "Código do responsável pelo setor"
         ToolTipText     =   "Código do responsável pelo setor"
         Top             =   480
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4683
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MAESTRO.chameleonButton cmdRetorno 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   11
         Tag             =   "Excluir treinamento"
         ToolTipText     =   "Excluir treinamento"
         Top             =   960
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
         MICON           =   "frmRetorna.frx":1056
         PICN            =   "frmRetorna.frx":1072
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdRetorno 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Tag             =   "Incluir treinamento"
         ToolTipText     =   "Incluir treinamento"
         Top             =   960
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
         MICON           =   "frmRetorna.frx":1D4C
         PICN            =   "frmRetorna.frx":1D68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmRetorna.frx":2A42
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRetorna.frx":2AAA
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin MAESTRO.chameleonButton cmdRetorno 
      Height          =   615
      Index           =   3
      Left            =   720
      TabIndex        =   14
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   6960
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
      MICON           =   "frmRetorna.frx":2B16
      PICN            =   "frmRetorna.frx":2B32
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdRetorno 
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   6960
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
      MICON           =   "frmRetorna.frx":380C
      PICN            =   "frmRetorna.frx":3828
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
Attribute VB_Name = "frmRetorna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsLocal As New ADODB.Recordset

Private Sub DTPicker1_Change()
    calculaDiasAfast
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub cmdRetorno_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirTreinamento
        LimpaControlesTreinamento
    Case 1
        mobjMsg.Abrir "Deseja EXCLUIR curso/treinamento das sugestões de retorno?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView1
            LimpaControlesTreinamento
        End If
    Case 2
        mobjMsg.Abrir "Confirma os dados de retorno de afastamento do colaborador?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            gravaDadosINTDPendCur
            Unload Me
        End If
    End Select
End Sub

Private Sub Command1_Click()
    ChamaGridCurso
    CarregaCurso
    CompoeComboNivel cboRetorno(5), txtRetorno(5)
End Sub

Private Sub Form_Load()
    listview_cabecalho
    Compoedados
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Compoedados()
    Dim rsRetornaColab As New ADODB.Recordset
    Dim SqlRetornaColab As String
    SqlRetornaColab = "select a.cpf,a.codcolaborador,a.nomecolaborador,a.dataafastamento,b.codmatriz,c.codcargo,d.nomecargo from tbcolaboradores  as a inner join tbcolaboradoreshist as b on  a.cpf = b.cpf and " & _
    "b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo where a.cpf = '" & varGlobal & "' and a.ativo = 'A'"
    rsRetornaColab.Open SqlRetornaColab, cnBanco, adOpenKeyset, adLockReadOnly
    DTPicker2 = rsRetornaColab.Fields(3)
    txtRetorno(0).Text = rsRetornaColab.Fields(1)
    txtRetorno(1).Text = rsRetornaColab.Fields(2)
    txtRetorno(2).Text = Format(rsRetornaColab.Fields(4), "0000") & " - " & rsRetornaColab.Fields(6)
    txtRetorno(3).Text = rsRetornaColab.Fields(0)
    calculaDiasAfast
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Nome curso/treinamento", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Nível", ListView1.Width / 5
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub calculaDiasAfast()
    SkinLabel7 = DTPicker1 - DTPicker2
    If Val(SkinLabel7) >= Val(vAfastDias) Then
        TravaDestravaTreinamentos True
        filtraLVTrei
    Else
        ListView1.ListItems.Clear
        TravaDestravaTreinamentos False
    End If
End Sub

Private Sub filtraLVTrei()
    Dim rsTrei As New ADODB.Recordset
    Dim sqlTrei As String
    Dim ItemLst As ListItem
    Dim X As Integer
    Dim vValidador As Boolean
    X = 0
    ''1º Monta os treinamentos INTRODUTORIOS da Matriz (EM Fase de Teste)
    If vAfastTreiInt = "S" Then
        sqlTrei = "select c.codmatriz,a.codtreinamento,c.nivel,b.nometreinamento,d.nomenivel from tbTreinamentosint as a inner join tbTreinamentos as b on a.codtreinamento = b.codtreinamento " & _
        "left join tbmatriz as c on a.codsetor = c.codsetor left join tbTreinamentosNiv as d on c.nivel = d.codnivel Where a.codsetor = 0 or c.codmatriz = '" & Val(Mid$(txtRetorno(2), 1, 6)) & "'"
        rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockReadOnly
    
        Y = ListView1.ListItems.Count
        If Y > 0 Then
            vValidador = True
            While Not rsTrei.EOF
                For X = 1 To Y
                    ListView1.ListItems.Item(X).Selected = True
                    If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(1) Then
                        vValidador = False
                    End If
                Next
                If vValidador = True Then
                        Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
                        ItemLst.SubItems(1) = "" & rsTrei.Fields(3)
                        If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") & " - " & rsTrei.Fields(5) Else ItemLst.SubItems(2) = "-"
                        'ItemLst.SubItems(3) = "-"
                        'ItemLst.SubItems(4) = "-"
                        'ItemLst.SubItems(5) = "-"
                End If
                rsTrei.MoveNext
                vValidador = True
                Y = ListView1.ListItems.Count
            Wend
        Else
            While Not rsTrei.EOF
                Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
                ItemLst.SubItems(1) = "" & rsTrei.Fields(3)
                If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") & " - " & rsTrei.Fields(5) Else ItemLst.SubItems(2) = "-"
                'ItemLst.SubItems(3) = "-"
                'ItemLst.SubItems(4) = "-"
                'ItemLst.SubItems(5) = "-"
                rsTrei.MoveNext
                'X = X + 1
            Wend
        End If
        rsTrei.Close
    End If
    
    ''--------------------------------------------------------------
    ''2º Monta os treinamentos Obrigatórios da Matriz (EM Fase de Teste)
    If vAfastTreiObr = "S" Then
        sqlTrei = "select c.codmatriz,a.codtreinamento,c.nivel,b.nometreinamento,d.nomenivel from tbTreinamentosobr as a inner join tbTreinamentos as b on a.codtreinamento = b.codtreinamento " & _
        "left join tbmatriz as c on a.codsetor = c.codsetor left join tbTreinamentosNiv as d on c.nivel = d.codnivel Where a.codsetor = 0 or c.codmatriz = '" & Val(Mid$(txtRetorno(2), 1, 6)) & "'"
        rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockOptimistic
    
        Y = ListView1.ListItems.Count
        If Y > 0 Then
            vValidador = True
            While Not rsTrei.EOF
                For X = 1 To Y
                    ListView1.ListItems.Item(X).Selected = True
                    If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(1) Then
                        vValidador = False
                    End If
                Next
                If vValidador = True Then
                        Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
                        ItemLst.SubItems(1) = "" & rsTrei.Fields(3)
                        If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") & " - " & rsTrei.Fields(5) Else ItemLst.SubItems(2) = "-"
                        'ItemLst.SubItems(3) = "-"
                        'ItemLst.SubItems(4) = "-"
                        'ItemLst.SubItems(5) = "-"
                End If
                rsTrei.MoveNext
                vValidador = True
                Y = ListView1.ListItems.Count
            Wend
        Else
            While Not rsTrei.EOF
                Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
                ItemLst.SubItems(1) = "" & rsTrei.Fields(3)
                If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") & " - " & rsTrei.Fields(5) Else ItemLst.SubItems(2) = "-"
                'ItemLst.SubItems(3) = "-"
                'ItemLst.SubItems(4) = "-"
                'ItemLst.SubItems(5) = "-"
                rsTrei.MoveNext
                'X = X + 1
            Wend
        End If
        rsTrei.Close
    End If
    Set rsTrei = Nothing
End Sub

Private Sub txtRetorno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCurso
            CompoeComboNivel cboRetorno(5), txtRetorno(5)
        End If
    End Select
End Sub

Private Sub CarregaCurso()
    Dim X As Integer
    Dim SqlCursos As String
    Dim rsCursos As New ADODB.Recordset
    SqlCursos = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by tbTreinamentos.codtreinamento"
    rsCursos.Open SqlCursos, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsCursos.EOF Then rsCursos.MoveFirst
    rsCursos.Find "codtreinamento=" & "'" & Val(Me.txtRetorno(5)) & "'"
    If rsCursos.EOF Then
        txtRetorno(5).Text = Format(txtRetorno(5), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Curso/Treinamento não cadastrado", Ok, critico, "Atenção"
            txtRetorno(6) = ""
        End If
    Else
        txtRetorno(5).Text = Format(rsCursos.Fields(0), "000000") & ""
        txtRetorno(6).Text = rsCursos.Fields(1)
    End If
    rsCursos.Close
    Set rsCursos = Nothing
End Sub

Private Sub ChamaGridCurso()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and introdutorio = 'N' order by tbTreinamentos.nometreinamento"
    procnom = "nometreinamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Treinamento"
    Pesquisa = frmINTD.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nometreinamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtRetorno(5).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub IncluirTreinamento()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaTreinamento = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.ListItems.Item(X) = Me.txtRetorno(5) Then
                Me.txtRetorno(5) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtRetorno(6)
                ListView1.SelectedItem.ListSubItems.Item(2) = cboRetorno(5)
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtRetorno(5))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtRetorno(5))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtRetorno(6)
    ItemLst.SubItems(2) = cboRetorno(5)
    txtRetorno(5).SetFocus
End Sub

Private Function ValidaTreinamento()
    ValidaTreinamento = False
    If txtRetorno(6).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtRetorno(6).Tag, Ok, critico, "Atenção"
        Me.txtRetorno(5).SetFocus
        Exit Function
    End If
    ValidaTreinamento = True
End Function

Private Sub LimpaControlesTreinamento()
    Dim X As Integer
    cboRetorno(5).Text = ""
    For X = 5 To 6
        txtRetorno(X) = ""
    Next
    txtRetorno(5).SetFocus
End Sub

Private Sub TravaDestravaTreinamentos(sIt As String)
    cboRetorno(5).Enabled = sIt
    Command1.Enabled = sIt
    For X = 5 To 6
        txtRetorno(X).Enabled = sIt
    Next
    cmdRetorno(0).Enabled = sIt
    cmdRetorno(1).Enabled = sIt
    ListView1.Enabled = sIt
End Sub

'ABAIXO DE GRAVAÇÃO DE TREINAMENTOS
Private Sub gravaDadosINTDPendCur()
    If Val(SkinLabel7) >= Val(vAfastDias) Then
        'Grava todos os treinamentos listador no form frmINTD na tabela
        'tbPendentesCur
        Dim rsTreiPen As New ADODB.Recordset
        Dim sqlTreiPen As String
        Dim rsDeletar As New ADODB.Recordset
        Dim sqlDeletar As String
        Dim contaID As Integer
    
        sqlTreiPen = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' order by id"
        rsTreiPen.Open sqlTreiPen, cnBanco, adOpenKeyset, adLockReadOnly
        If Not rsTreiPen.EOF Then
            rsTreiPen.MoveLast
            contaID = rsTreiPen.Fields(5) + 1
        Else
            contaID = 1
        End If
        rsTreiPen.Close
        Set rsTreiPen = Nothing
    
        sqlTreiPen = "Select * from tbPendentesCur as a where codcoligada = '" & vCodcoligada & "'"
        rsTreiPen.Open sqlTreiPen, cnBanco, adOpenKeyset, adLockOptimistic
        Dim ItemLst As ListItem
        Dim X As Integer, Y As Integer
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            rsTreiPen.AddNew
            rsTreiPen.Fields(0) = txtRetorno(3) 'CPF
            rsTreiPen.Fields(1) = Val(Mid$(txtRetorno(2), 1, 6)) 'Codigo da matriz
            rsTreiPen.Fields(2) = Val(ListView1.ListItems.Item(X))
            rsTreiPen.Fields(4) = "S"
            rsTreiPen.Fields(5) = contaID
            rsTreiPen.Fields(6) = "Pendente"
            rsTreiPen.Fields(7) = 0
            If ListView1.SelectedItem.ListSubItems.Item(2) <> "-" Then
                rsTreiPen.Fields(12) = Val(ListView1.SelectedItem.ListSubItems.Item(2))
            Else
                rsTreiPen.Fields(12) = 0
            End If
            'rsTreiPen.Fields(13) = Val(txtINTD(0))
            rsTreiPen.Fields(14) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        Next
        If Y > 0 Then rsTreiPen.Update
        rsTreiPen.Close
        Set rsTreiPen = Nothing
    End If

    Dim rsAfastColab As New ADODB.Recordset
    Dim SqlAfastColab As String

    SqlAfastColab = "Update tbcolaboradores set Ativo = CASE WHEN Ativo = 'S' then 'A' WHEN Ativo = 'A' then 'S' END where cpf = '" & varGlobal & "';" & _
    "Update tbcolaboradores set dataafastamento = null where cpf = '" & varGlobal & "'"
    rsAfastColab.Open SqlAfastColab, cnBanco
    mobjMsg.Abrir "Colaborador retornou do afastamento", Ok, informacao, "SGC"
End Sub
