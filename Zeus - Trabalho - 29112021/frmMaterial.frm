VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmMaterial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Material"
   ClientHeight    =   8625
   ClientLeft      =   1755
   ClientTop       =   705
   ClientWidth     =   9960
   Icon            =   "frmMaterial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   3720
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.CommandButton chameleonButton1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      Picture         =   "frmMaterial.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "Importar Fórmula"
      ToolTipText     =   "Importar Fórmula"
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton chamCad 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   720
      Picture         =   "frmMaterial.frx":1CD2B
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton chamCad 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      Picture         =   "frmMaterial.frx":1D9F5
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   7920
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Fórmulas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   9735
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   8520
         TabIndex        =   24
         Tag             =   "Constante para cálculo de área de pintura"
         ToolTipText     =   "Constante para cálculo de área de pintura"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Tag             =   "Fórmula para calculo de peso"
         ToolTipText     =   "Fórmula para calculo de peso"
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   4440
         TabIndex        =   22
         Tag             =   "Fórmula para calculo de área de pintura"
         ToolTipText     =   "Fórmula para calculo de área de pintura"
         Top             =   480
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   8520
         OleObjectBlob   =   "frmMaterial.frx":1E6BF
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "frmMaterial.frx":1E733
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":1E79B
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Constantes "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   5535
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
         Height          =   345
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "Constante da fórmula"
         ToolTipText     =   "Constante da fórmula"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton chamCad 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   1320
         Picture         =   "frmMaterial.frx":1E7FD
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "Exclui constante selecionada"
         ToolTipText     =   "Exclui constante selecionada"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton chamCad 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmMaterial.frx":1F4C7
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "Editar constante"
         ToolTipText     =   "Editar constante"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton chamCad 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmMaterial.frx":20191
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "Insere nova constante"
         ToolTipText     =   "Insere nova constante"
         Top             =   1800
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   735
         Left            =   2040
         OleObjectBlob   =   "frmMaterial.frx":20E5B
         TabIndex        =   16
         Top             =   1680
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmMaterial.frx":20F4F
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Tag             =   "Descrição da constante"
         ToolTipText     =   "Descrição da constante"
         Top             =   1080
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":20FBD
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":21039
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
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
   Begin VB.Frame Frame3 
      Caption         =   "Informações Gerais "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5760
      TabIndex        =   4
      Top             =   3000
      Width           =   4095
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "Informações Gerais"
         ToolTipText     =   "Informações Gerais"
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Material "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":210B7
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":21131
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2760
         TabIndex        =   1
         Tag             =   "Descrição"
         ToolTipText     =   "Descrição"
         Top             =   480
         Width           =   6855
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Tag             =   "Código do Material"
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmMaterial.frx":211A5
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmMaterial.frx":21211
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   7800
         TabIndex        =   3
         Top             =   7185
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsMaterial As New ADODB.Recordset
Private sqlMaterial As String
Private Status As String
Private rsLocal As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados do Material?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravarDados
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Materiais?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Unload Me
            Set frmMaterial = Nothing
        End If
    End Select
End Sub

Private Sub cmdCadastro_MouseOver(Index As Integer)
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub chameleonButton1_Click()
    ChamaGridFormula
End Sub

Private Sub ChamaGridFormula()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "Select a.idprd,b.CODIGOPRD,b.NOMEFANTASIA from tbMateriais as a inner join " & vBancoTotvs & ".dbo.TPRD as b on a.idprd = b.idprd order by b.nomefantasia"
    procnom = "nomefantasia"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Produtos com fórmulas"
    Pesquisa = frmMaterial.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "CODIGOPRD=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            'Text5.Text = rsLocal.Fields(0)
            Pesquisa = rsLocal.Fields(0)
            importaFormula Val(Pesquisa)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
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

Private Sub importaFormula(vID As Integer)
On Error GoTo Err
    Dim rsimportaFormula As New ADODB.Recordset
    Dim sqlimportaFormula As String
    sqlimportaFormula = "Select a.idprd,a.formula,a.forpint,a.constpint,b.valconst,b.descricao,b.idseq,a.observacao from tbMateriais as a inner join tbConstantes as b on a.idprd = b.idprd where a.idprd = '" & vID & "' order by b.idseq"
    rsimportaFormula.Open sqlimportaFormula, cnBanco, adOpenKeyset, adLockReadOnly
    txtcadastro(2).Text = rsimportaFormula.Fields(1) 'Formula de peso
    txtcadastro(6).Text = rsimportaFormula.Fields(2) 'Formula de pintura
    txtcadastro(3).Text = rsimportaFormula.Fields(3) 'Constante de pintura
    txtcadastro(7).Text = rsimportaFormula.Fields(7) 'Observação
    ListView1.ListItems.Clear
    While Not rsimportaFormula.EOF
        Set ItemLst = ListView1.ListItems.Add(, , "const0(" & rsimportaFormula.Fields(6) & ")")
        ItemLst.SubItems(1) = "" & Format(rsimportaFormula.Fields(4), "#,##0.000000000;(#,##0.000000000)")
        ItemLst.SubItems(2) = "" & rsimportaFormula.Fields(5)
        rsimportaFormula.MoveNext
    Wend
    Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
    rsimportaFormula.Close
    Set rsimportaFormula = Nothing
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
'    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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
        Compoe_Listview
        'DesbloqueiaControles
    End If
    'configControles
    carregarIconBotao
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao chamCad(0), 0, 46 'Inserir
    carregaImagemBotao chamCad(1), 1, 32 'Editar
    carregaImagemBotao chamCad(2), 2, 33 'Excluir
    carregaImagemBotao chamCad(3), 3, 45 'Salvar
    carregaImagemBotao chamCad(4), 4, 34 'Sair
    carregaImagemBotao chameleonButton1, 4, 54 'Importar Fórmula
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Add , , "Nome Const", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Valor constante", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Descricao da Constante", ListView1.Width / 2
    ListView1.View = lvwReport
End Sub

Private Sub Compoe_Listview()
On Error GoTo Err
    Dim rsLisview As New ADODB.Recordset
    Dim sql As String
    Dim ItemLst As ListItem
    sql = "select * from tbconstantes Where tbconstantes.idprd= '" & Val(txtcadastro(9)) & "' order by idseq"
    rsLisview.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
    ListView1.ListItems.Clear
    While Not rsLisview.EOF
        'insere o item do arquivo de dados
        Set ItemLst = ListView1.ListItems.Add(, , "const0(" & rsLisview.Fields(3) & ")")
        'cada item precisa de um subitem para exibir na lista
        ItemLst.SubItems(1) = "" & Format(rsLisview.Fields(1), "#,##0.000000000;(#,##0.000000000)")
        '#,##0.000000000;(#,##0.000000000)
        ItemLst.SubItems(2) = "" & rsLisview.Fields(2)
        'vai para o proximo registro
        rsLisview.MoveNext
    Wend
    Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
    rsLisview.Close
    ListView1.Refresh
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

Private Sub GravarDados()
On Error GoTo Err
    Dim rsSalvar As New ADODB.Recordset
    Dim rsSalvaMat As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim y As Integer
    If ValidaGravacao = False Then Exit Sub
10  cnBanco.BeginTrans
    SqlSalvar = "Delete from tbconstantes where tbconstantes.idprd= " & Val(Me.txtcadastro(9))
    rsSalvar.Open SqlSalvar, cnBanco

    SqlSalvar = "Select * from tbconstantes"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For x = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(x).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtcadastro(9))
        rsSalvar.Fields(1) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(3) = x
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    
    SqlSalvar = "select * from tbmateriais where tbmateriais.idprd = '" & txtcadastro(9) & "'"
    rsSalvaMat.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvaMat.EOF Then rsSalvaMat.AddNew
    rsSalvaMat.Fields(0) = Val(txtcadastro(9))
    rsSalvaMat.Fields(1) = txtcadastro(2)
    rsSalvaMat.Fields(2) = txtcadastro(3)
    rsSalvaMat.Fields(3) = txtcadastro(6)
    rsSalvaMat.Fields(4) = txtcadastro(7)
    rsSalvaMat.Update
    cnBanco.CommitTrans
    rsSalvar.Close
    Set rsSalvar = Nothing
    rsSalvaMat.Close
    Set rsSalvaMat = Nothing
    
    AtualizaListview
    mobjMsg.Abrir "Os dados do Material foram salvos com sucesso", Ok, informacao, "ZEUS"
    Unload Me
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub LimpaControles()
    Dim x As Integer
    DesbloqueiaControles
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x) = ""
    Next
    ListView1.ListItems.Clear
    txtcadastro(0).Text = Format(GeraCodigo, "000000")
    txtcadastro(0).Enabled = False
    For x = 1 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = True
    Next
    cboCadastro = ""
    ListView1.Enabled = True
End Sub

Private Sub CompoeControles()
    Dim x As Integer
    txtcadastro(9).Text = Format(rsMaterial.Fields(0), "000000")
    txtcadastro(0).Text = rsMaterial.Fields(1)
    txtcadastro(1).Text = rsMaterial.Fields(2)
    If Not IsNull(rsMaterial.Fields(4)) Then txtcadastro(8).Text = rsMaterial.Fields(4)
    If Not IsNull(rsMaterial.Fields(5)) Then txtcadastro(2).Text = rsMaterial.Fields(5)
    If Not IsNull(rsMaterial.Fields(6)) Then txtcadastro(6).Text = rsMaterial.Fields(6)
    If Not IsNull(rsMaterial.Fields(7)) Then txtcadastro(3).Text = rsMaterial.Fields(7)
    If Not IsNull(rsMaterial.Fields(8)) Then txtcadastro(7).Text = rsMaterial.Fields(8)
    'cbocadastro.Text = rsMaterial(4)
    'BloqueiaControles
End Sub

Private Function ValidaGravacao()
    ValidaGravacao = False
    If txtcadastro(2).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(2).Tag, vbInformation, "Atenção"
        Me.txtcadastro(2).SetFocus
        Exit Function
    End If
    If txtcadastro(6).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(6).Tag, vbInformation, "Atenção"
        Me.txtcadastro(6).SetFocus
        Exit Function
    End If
    If txtcadastro(3).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(3).Tag, vbInformation, "Atenção"
        Me.txtcadastro(3).SetFocus
        Exit Function
    End If
    ValidaGravacao = True
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If txtcadastro(0).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(0).Tag, vbInformation, "Atenção"
        Me.txtcadastro(x).SetFocus
        Exit Function
    End If
    If txtcadastro(1).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(1).Tag, vbInformation, "Atenção"
        Me.txtcadastro(x).SetFocus
        Exit Function
    End If
    If txtcadastro(2).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(2).Tag, vbInformation, "Atenção"
        Me.txtcadastro(x).SetFocus
        Exit Function
    End If
    If txtcadastro(6).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(6).Tag, vbInformation, "Atenção"
        Me.txtcadastro(x).SetFocus
        Exit Function
    End If
    'If cbocadastro.Text = "" Then
    '    Msgbox "Favor informar o campo " & Me.cbocadastro.Tag, vbInformation, "Atenção"
    '    Me.txtcadastro(X).SetFocus
    '    Exit Function
    'End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = False
    Next
    cboCadastro.Enabled = False
    ListView1.Enabled = False
End Sub

Private Sub DesbloqueiaControles()
    For x = 1 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = True
    Next
    'cbocadastro.Enabled = True
    ListView1.Enabled = True
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirMaterial
    SqlGera = "Select top 1 * from tbMateriais order by CODIGOPRD Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsMaterial.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtcadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharMaterial
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Private Sub AbrirMaterial()
On Error GoTo Err
    SqlM = "Select * from tbMateriais Order by codigoprd"
    rsMaterial.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
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

Private Sub FecharMaterial()
    rsMaterial.Close
    Set rsMaterial = Nothing
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    sqlMaterial = "select a.idprd,a.CODIGOPRD,a.NOMEFANTASIA,a.codtb2fat,c.descricao,b.formula,b.forpint,b.constpint,b.observacao from " & vBancoTotvs & ".dbo.TPRD as a left join " & sDatabaseName & ".dbo.tbMateriais as b on a.IDPRD = b.IDPRD left join " & vBancoTotvs & ".dbo.ttb2 as c on a.CODTB2FAT = c.CODTB2FAT where a.IDPRD = '" & Val(varGlobal) & "'"
    rsMaterial.Open sqlMaterial, cnBanco, adOpenKeyset, adLockReadOnly
    If rsMaterial.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Código do Material não encontrado", Ok, critico, "Atenção"
    End If
    rsMaterial.Close
    Set rsMaterial = Nothing
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

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    y = vListViewPrincipal.ListItems.Count
    For x = 1 To y
        If vListViewPrincipal.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    vListViewPrincipal.SelectedItem.ListSubItems.Item(5) = txtcadastro(2).Text
    vListViewPrincipal.SelectedItem.ListSubItems.Item(6) = txtcadastro(6).Text
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub chamCad_Click(Index As Integer)
    Select Case Index
    Case 0
        ListView1.Enabled = True
        IncluirItem
        txtcadastro(4).SetFocus
    Case 1
        AlterarItem
    Case 2
        ExcluirItem
    Case 3
        CancelaSN = 1
        GravarDados
        AtualizaListview
        Unload Me
    Case 4
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Materiais?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            CancelaSN = 0
            Unload Me
        End If
    End Select
End Sub

Private Sub IncluirItem()
    Dim ItemLst As ListItem
    Dim x As Integer, y As Integer
    If ValidaCampo = False Then Exit Sub
    y = ListView1.ListItems.Count
    If y > 0 Then
        For x = 1 To y
            If ListView1.ListItems.Item(x) = Me.Text1 Then
                Me.Text1 = ListView1.ListItems.Item(x)
                ListView1.SelectedItem.ListSubItems.Item(1) = Format(txtcadastro(4), "###,##0.000000000;(###,##0.000000000)")
                ListView1.SelectedItem.ListSubItems.Item(2) = Me.txtcadastro(5).Text
                Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
                
                txtcadastro(4) = ""
                txtcadastro(5) = ""
                Me.Text1 = ""
                y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , "const0(" & ListView1.ListItems.Count + 1 & ")")
        y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , "const0(" & ListView1.ListItems.Count + 1 & ")")
        y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = Format(txtcadastro(4), "###,##0.000000000;(###,##0.000000000)")
    ItemLst.SubItems(2) = Me.txtcadastro(5).Text
    Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
    txtcadastro(4) = ""
    txtcadastro(5) = ""
    Me.Text1 = ""
End Sub

Private Sub ExcluirItem()
    Dim x As Integer, y As Integer
    y = ListView1.ListItems.Count
    If y = 0 Then Exit Sub
    For x = 1 To y
        If ListView1.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    ListView1.ListItems.Remove (x)
End Sub

Private Sub AlterarItem()
    Dim y As Integer, x As Integer
    y = ListView1.ListItems.Count
    For x = 1 To y
        If ListView1.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    Me.Text1.Text = ListView1.ListItems.Item(x)
    Me.txtcadastro(4).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.txtcadastro(5).Text = ListView1.SelectedItem.ListSubItems.Item(2)
End Sub

Private Sub ListView1_DblClick()
    AlterarItem
End Sub

