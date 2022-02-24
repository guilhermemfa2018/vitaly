VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDepartamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de departamentos"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmDepartamentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   6
      Left            =   720
      TabIndex        =   18
      Top             =   7200
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
      MICON           =   "frmDepartamentos.frx":0CCA
      PICN            =   "frmDepartamentos.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   7200
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
      MICON           =   "frmDepartamentos.frx":19C0
      PICN            =   "frmDepartamentos.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "Status do departamento"
         ToolTipText     =   "Status do departamento"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Histórico de responsáveis pelo departamento "
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   8295
      Begin VB.CommandButton cmdteste 
         Caption         =   "..."
         Height          =   255
         Left            =   7800
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Tag             =   "Data fim das atividades de responsabilidade pelo departamento"
         ToolTipText     =   "Data fim das atividades de responsabilidade pelo departamento"
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   57278467
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Tag             =   "Data início das atividades de responsabilidade pelo departamento"
         ToolTipText     =   "Data início das atividades de responsabilidade pelo departamento"
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   40498
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label6 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmDepartamentos.frx":26B6
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDepartamentos.frx":2726
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCadDepartamento 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   6135
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label4 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmDepartamentos.frx":279C
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCadDepartamento 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Tag             =   "Código do responsável pelo departamento"
         ToolTipText     =   "Código do responsável pelo departamento"
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDepartamentos.frx":2804
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   16
         Top             =   1560
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
         MICON           =   "frmDepartamentos.frx":2870
         PICN            =   "frmDepartamentos.frx":288C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   15
         Top             =   1560
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
         MICON           =   "frmDepartamentos.frx":3566
         PICN            =   "frmDepartamentos.frx":3582
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   1560
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
         MICON           =   "frmDepartamentos.frx":425C
         PICN            =   "frmDepartamentos.frx":4278
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1560
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
         MICON           =   "frmDepartamentos.frx":4F52
         PICN            =   "frmDepartamentos.frx":4F6E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   7
         Tag             =   "Histórico dos reponsávei pelo departamento"
         ToolTipText     =   "Histórico dos reponsávei pelo departamento"
         Top             =   2280
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4048
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do departamento"
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtCadDepartamento 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do departamento"
         ToolTipText     =   "Nome do departamento"
         Top             =   480
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmDepartamentos.frx":5C48
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCadDepartamento 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do departamento"
         ToolTipText     =   "Código do departamento"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDepartamentos.frx":5CD0
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCadDepartamento 
         Height          =   975
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "Breve descrição do departamento"
         ToolTipText     =   "Breve descrição do departamento"
         Top             =   1080
         Width           =   8055
      End
      Begin VB.Label Label7 
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDepartamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsDepartamentos As New ADODB.Recordset
Private SqlDepartamentos As String
Private rsColaborador As New ADODB.Recordset
Private SqlColaborador As String
Private rsHistorico As New ADODB.Recordset
Private SqlHistorico As String
Private rsLocal As New ADODB.Recordset
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirResponsavel
        LimpaControlesResp
    Case 1
        LimpaControlesResp
    Case 2
        AlteraResponsavel
    Case 3
        ExcluirItemLV ListView1
    'Case 4
        'ChamaGridColaborador
        'CarregaColaborador
    Case 5
        mobjMsg.Abrir "Deseja salvar os dados do departamento?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            GravarDados
            gravaLog "Código dep.: " & txtCadDepartamento(0), "Nome dep: " & txtCadDepartamento(1), ""
            'AtivaLD
            Pesquisa = 0
            Unload Me
            Set frmDepartamentos = Nothing
        End If
    Case 6
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Departamentos?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Pesquisa = 0
            Unload Me
        End If
    End Select
End Sub

Private Sub cmdCadastro_MouseOver(Index As Integer)
    Legenda = cmdCadastro(Index).ToolTipText
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdteste_Click()
    ChamaGridColaborador
    CarregaColaborador
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    listview_cabecalho
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        'DesbloqueiaControles
    End If
    configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    'OrganizaForm
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Registro", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Responsável", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "Data Início", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Data Fim", ListView1.Width / 6
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarDepartamento As New ADODB.Recordset
    Dim SqlSalvarDepartamento As String
    Dim rsDeletarHistResp As New ADODB.Recordset
    Dim SqlDeletarHistResp As String
    Dim rsSalvarHistResp As New ADODB.Recordset
    Dim SqlSalvarHistResp As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarDepartamento = "select * from tbDepartamentos where codcoligada = '" & vCodcoligada & "' and coddepartamento = '" & txtCadDepartamento(0) & "'"
    rsSalvarDepartamento.Open SqlSalvarDepartamento, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarDepartamento.EOF Then rsSalvarDepartamento.AddNew
    rsSalvarDepartamento.Fields(0) = Val(txtCadDepartamento(0))
    rsSalvarDepartamento.Fields(1) = txtCadDepartamento(1)
    rsSalvarDepartamento.Fields(2) = txtCadDepartamento(2)
    If Check1.Value = 0 Then
        rsSalvarDepartamento.Fields(3) = "N"
    Else
        rsSalvarDepartamento.Fields(3) = "S"
    End If
    rsSalvarDepartamento.Fields(4) = vCodcoligada 'Codigo da coligada
    
    rsSalvarDepartamento.Update
    
    SqlDeletarHistResp = "Delete from tbDepartamentosHistResp where codcoligada ='" & vCodcoligada & "' and coddepartamento = '" & Val(txtCadDepartamento(0)) & "'"
    rsDeletarHistResp.Open SqlDeletarHistResp, cnBanco
      
    SqlSalvarHistResp = "select * from tbDepartamentosHistResp"
    rsSalvarHistResp.Open SqlSalvarHistResp, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvarHistResp.AddNew
        rsSalvarHistResp.Fields(0) = Val(txtCadDepartamento(0))
        rsSalvarHistResp.Fields(1) = ListView1.ListItems.Item(X)
        rsSalvarHistResp.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
        If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then rsSalvarHistResp.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(3)
        rsSalvarHistResp.Fields(4) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvarHistResp.EOF Then rsSalvarHistResp.Update
    cnBanco.CommitTrans
    rsSalvarDepartamento.Close
    Set rsSalvarDepartamento = Nothing
    mobjMsg.Abrir "Os dados do Departamento foram salvos com sucesso", Ok, informacao, "SGC"
    AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    'DesbloqueiaControles
    For X = 0 To txtCadDepartamento.Count - 1
        txtCadDepartamento(X) = ""
    Next
    txtCadDepartamento(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub LimpaControlesResp()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker2.Value = ""
    txtCadDepartamento(3) = ""
    txtCadDepartamento(4) = ""
    txtCadDepartamento(3).Enabled = True
    txtCadDepartamento(4).Enabled = True
    cmdCadastro(4).Enabled = True
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadDepartamento(0).Text = Format(rsDepartamentos.Fields(0), "000000")
    txtCadDepartamento(1).Text = rsDepartamentos.Fields(1)
    If Not IsNull(rsDepartamentos.Fields(2)) Then txtCadDepartamento(2).Text = rsDepartamentos.Fields(2)
    If rsDepartamentos.Fields(3) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Sub Compoe_Listview()
    ' Declaração de variaveis
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsHistorico.EOF
        Set ItemLst = ListView1.ListItems.Add(, , rsHistorico.Fields(1))
        ItemLst.SubItems(1) = "" & rsHistorico.Fields(5)
        ItemLst.SubItems(2) = "" & rsHistorico.Fields(2)
        If Not IsNull(rsHistorico.Fields(3)) Then ItemLst.SubItems(3) = rsHistorico.Fields(3) Else ItemLst.SubItems(3) = "-"
        rsHistorico.MoveNext
        X = X + 1
    Wend
    'Legenda = ""
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadDepartamento(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadDepartamento(0).Tag, Ok, critico, "Atenção"
        Me.txtCadDepartamento(0).SetFocus
        Exit Function
    End If
    If txtCadDepartamento(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadDepartamento(1).Tag, Ok, critico, "Atenção"
        Me.txtCadDepartamento(1).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadDepartamento.Count - 1
        txtCadDepartamento(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadDepartamento.Count - 1
        txtCadDepartamento(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirDepartamento
    SqlGera = "Select top 1 * from tbDepartamentos where codcoligada = '" & vCodcoligada & "' order by codDepartamento Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsDepartamentos.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadDepartamento(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharDepartamentos
End Function

Private Sub AbrirDepartamento()
    SqlDepartamentos = "Select * from tbDepartamentos where codcoligada = '" & vCodcoligada & "' Order by codDepartamento"
    rsDepartamentos.Open SqlDepartamentos, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharDepartamentos()
    rsDepartamentos.Close
    Set rsDepartamentos = Nothing
End Sub

Private Sub AbrirHistorico()
    SqlHistorico = "Select tbDepartamentosHistResp.*,tbcolaboradores.nomecolaborador from tbDepartamentosHistResp inner join tbcolaboradores on tbcolaboradores.codcolaborador = tbDepartamentosHistResp.codcolaborador where tbDepartamentosHistResp.codcoligada = '" & vCodcoligada & "' and coddepartamento = '" & Val(txtCadDepartamento(0)) & "'"
    rsHistorico.Open SqlHistorico, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharHistorico()
    rsHistorico.Close
    Set rsHistorico = Nothing
End Sub

Private Sub ResultPesq()
    SqlDepartamentos = "Select * from tbDepartamentos Where tbDepartamentos.codcoligada = '" & vCodcoligada & "' and tbDepartamentos.codDepartamento= '" & Val(varGlobal) & "' order by codDepartamento"
    rsDepartamentos.Open SqlDepartamentos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsDepartamentos.RecordCount > 0 Then
        CompoeControles
        AbrirHistorico
        Compoe_Listview
        FecharHistorico
    Else
        mobjMsg.Abrir "Departamento não encontrado", Ok, critico, "Atenção"
    End If
    rsDepartamentos.Close
    Set rsDepartamentos = Nothing
End Sub

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadDepartamento(0), "000000"))
        ItemLst.SubItems(1) = txtCadDepartamento(1).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadDepartamento(1).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub CarregaColaborador()
    Dim X As Integer
    SqlColaborador = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' and ativo = 'S' order by codcolaborador"
    rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsColaborador.EOF Then rsColaborador.MoveFirst
    rsColaborador.Find "codcolaborador=" & "'" & Me.txtCadDepartamento(3) & "'"
    If rsColaborador.EOF Then
        txtCadDepartamento(3).Text = txtCadDepartamento(3)
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Colaborador não cadastrado", Ok, critico, "Atenção"
            txtCadDepartamento(4) = ""
        End If
    Else
        txtCadDepartamento(3).Text = rsColaborador.Fields(1)
        txtCadDepartamento(4).Text = rsColaborador.Fields(3)
    End If
    rsColaborador.Close
    Set rsColaborador = Nothing
End Sub

Private Sub ChamaGridColaborador()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' and ativo = 'S' order by nomecolaborador"
    procnom = "nomecolaborador"
    campo = 3
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Colaborador"
    Pesquisa = frmDepartamentos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecolaborador=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadDepartamento(3).Text = rsLocal.Fields(1)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

' (INICIO) >>>>>>>> CONTROLES DOS BOTOES DE HISTÓRICO DE RESPONSÁVEIS DO DEPARTAMENTO<<<<<<<<<<
Private Sub IncluirResponsavel()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaResponsavel = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtCadDepartamento(3) Then
                Me.txtCadDepartamento(3) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadDepartamento(4)
                ListView1.SelectedItem.ListSubItems.Item(2) = DTPicker1
                If DTPicker2.Value <> "" Then ListView1.SelectedItem.ListSubItems.Item(3) = DTPicker2 Else ListView1.SelectedItem.ListSubItems.Item(3) = "-"
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtCadDepartamento(3))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtCadDepartamento(3))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadDepartamento(4)
    ItemLst.SubItems(2) = DTPicker1
    If DTPicker2.Value <> "" Then ItemLst.SubItems(3) = DTPicker2 Else ItemLst.SubItems(3) = "-"
    txtCadDepartamento(3).SetFocus
End Sub

Private Function ValidaResponsavel()
    ValidaResponsavel = False
    If txtCadDepartamento(3).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadDepartamento(3).Tag, Ok, critico, "Atenção"
        Me.txtCadDepartamento(3).SetFocus
        Exit Function
    End If
    ValidaResponsavel = True
End Function

Private Sub AlteraResponsavel()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadDepartamento(3).Text = ListView1.ListItems.Item(X)
    Me.txtCadDepartamento(4).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    DTPicker1 = ListView1.SelectedItem.ListSubItems.Item(2)
    If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then DTPicker2 = ListView1.SelectedItem.ListSubItems.Item(3)
    txtCadDepartamento(3).Enabled = False
    txtCadDepartamento(4).Enabled = False
    txtCadDepartamento(4).Enabled = False
End Sub
' (FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraResponsavel
    End If
End Sub

Private Sub txtCadDepartamento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Error
    Select Case Index
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    End If
    If vExc = "N" Then
        cmdCadastro(3).UseGreyscale = True
        cmdCadastro(3).DragMode = 1
        cmdCadastro(3).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then
        cmdCadastro(5).UseGreyscale = True
        cmdCadastro(5).DragMode = 1
        cmdCadastro(5).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub Form_Resize()
    'OrganizaControles
End Sub

Private Function OrganizaForm()
    Me.Move 0, 0, Principal.ScaleWidth - 200, Principal.ScaleHeight - 350
End Function

Private Function OrganizaControles()
    On Error Resume Next
    Frame1.Move 1920, 0, Me.ScaleWidth - 2000
    Frame3.Move 0, 0, Me.ScaleWidth - Me.ScaleWidth + 1800, Me.ScaleHeight
    DBGrid2.Move 1920, 1080, Me.ScaleWidth - 2000, Me.ScaleHeight - 1080
End Function

