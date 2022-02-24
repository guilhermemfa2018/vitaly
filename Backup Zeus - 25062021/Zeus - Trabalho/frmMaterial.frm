VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin ZEUS.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   9240
      TabIndex        =   33
      Tag             =   "Importar Fórmula"
      ToolTipText     =   "Importar Fórmula"
      Top             =   7920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   11
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
      MICON           =   "frmMaterial.frx":0CCA
      PICN            =   "frmMaterial.frx":0CE6
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
      Caption         =   "Fórmulas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   9735
      Begin VB.TextBox txtcadastro 
         Height          =   285
         Index           =   3
         Left            =   8520
         TabIndex        =   29
         Tag             =   "Constante para cálculo de área de pintura"
         ToolTipText     =   "Constante para cálculo de área de pintura"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtcadastro 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Tag             =   "Fórmula para calculo de peso"
         ToolTipText     =   "Fórmula para calculo de peso"
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtcadastro 
         Height          =   285
         Index           =   6
         Left            =   4440
         TabIndex        =   27
         Tag             =   "Fórmula para calculo de área de pintura"
         ToolTipText     =   "Fórmula para calculo de área de pintura"
         Top             =   480
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   8520
         OleObjectBlob   =   "frmMaterial.frx":19C0
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "frmMaterial.frx":1A3A
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":1AA8
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Constantes "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   5535
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   495
         Left            =   2280
         OleObjectBlob   =   "frmMaterial.frx":1B10
         TabIndex        =   21
         Top             =   1800
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frmMaterial.frx":1C0A
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtcadastro 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Tag             =   "Descrição da constante"
         ToolTipText     =   "Descrição da constante"
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtcadastro 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Tag             =   "Constante da fórmula"
         ToolTipText     =   "Constante da fórmula"
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":1C78
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":1CFA
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Tag             =   "Exclui constante selecionada"
         ToolTipText     =   "Exclui constante selecionada"
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         MICON           =   "frmMaterial.frx":1D7E
         PICN            =   "frmMaterial.frx":1D9A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Tag             =   "Editar constante"
         ToolTipText     =   "Editar constante"
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         MICON           =   "frmMaterial.frx":2A74
         PICN            =   "frmMaterial.frx":2A90
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Tag             =   "Insere nova constante"
         ToolTipText     =   "Insere nova constante"
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         MICON           =   "frmMaterial.frx":376A
         PICN            =   "frmMaterial.frx":3786
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
         TabIndex        =   15
         Top             =   2400
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4048
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
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações Gerais "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5760
      TabIndex        =   5
      Top             =   3000
      Width           =   4095
      Begin VB.TextBox txtcadastro 
         Height          =   4455
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "Informações Gerais"
         ToolTipText     =   "Informações Gerais"
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Material "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":4460
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMaterial.frx":44E0
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   285
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
         Height          =   285
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
         OleObjectBlob   =   "frmMaterial.frx":455A
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmMaterial.frx":45CC
         TabIndex        =   16
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
   Begin ZEUS.chameleonButton chamCad 
      Height          =   615
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   11
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
      MICON           =   "frmMaterial.frx":4638
      PICN            =   "frmMaterial.frx":4654
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton chamCad 
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   7920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   11
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
      MICON           =   "frmMaterial.frx":532E
      PICN            =   "frmMaterial.frx":534A
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

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub chameleonButton1_Click()
    ChamaGridFormula
End Sub

Private Sub ChamaGridFormula()
    Dim F As New frmPesqger2
    'Dim Iposicao As Variant
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
End Sub

'Private Sub ChamaGridFormula()
'    Dim F As New frmpesqger
'    Sqlp = "Select a.idprd,b.CODIGOPRD,b.NOMEFANTASIA from tbMateriais as a inner join " & vBancoTotvs & ".dbo.TPRD as b on a.idprd = b.idprd order by b.nomefantasia"
'    procnom = "nomefantasia"
'    campo = 2
'    Campo1 = 1
'    campo2 = 2
'    Load F
'    F.Caption = "Produtos com fórmulas"
'    Pesquisa = frmMaterial.Tag
'    F.Show 1
'    If Pesquisa <> "" Then
'        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
'        If rsLocal.RecordCount < 1 Then Exit Sub
'        rsLocal.MoveFirst
'        rsLocal.Find "CODIGOPRD=" & "'" & Pesquisa & "'"
'        If Not rsLocal.EOF Then
'            Pesquisa = rsLocal.Fields(0)
'            importaFormula Val(Pesquisa)
'        End If
'        rsLocal.Close
'        Set rsLocal = Nothing
'    End If
'End Sub

Private Sub importaFormula(vID As Integer)
    Dim rsimportaFormula As New ADODB.Recordset
    Dim sqlimportaFormula As String
    sqlimportaFormula = "Select a.idprd,a.formula,a.forpint,a.constpint,b.valconst,b.descricao,b.idseq,a.observacao from tbMateriais as a inner join tbConstantes as b on a.idprd = b.idprd where a.idprd = '" & vID & "' order by b.idseq"
    rsimportaFormula.Open sqlimportaFormula, cnBanco, adOpenKeyset, adLockReadOnly
    txtCadastro(2).Text = rsimportaFormula.Fields(1) 'Formula de peso
    txtCadastro(6).Text = rsimportaFormula.Fields(2) 'Formula de pintura
    txtCadastro(3).Text = rsimportaFormula.Fields(3) 'Constante de pintura
    txtCadastro(7).Text = rsimportaFormula.Fields(7) 'Observação
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
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
'    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
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
        Compoe_Listview
        'DesbloqueiaControles
    End If
    'configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Add , , "Nome Const", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Valor constante", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Descricao da Constante", ListView1.Width / 2
    ListView1.View = lvwReport
End Sub

Private Sub Compoe_Listview()
    Dim rsLisview As New ADODB.Recordset
    Dim sql As String
    Dim ItemLst As ListItem
    sql = "select * from tbconstantes Where tbconstantes.idprd= '" & Val(txtCadastro(9)) & "' order by idseq"
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
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    Dim rsSalvar As New ADODB.Recordset
    Dim rsSalvaMat As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer
    If ValidaGravacao = False Then Exit Sub
    cnBanco.BeginTrans
    SqlSalvar = "Delete from tbconstantes where tbconstantes.idprd= " & Val(Me.txtCadastro(9))
    rsSalvar.Open SqlSalvar, cnBanco

    SqlSalvar = "Select * from tbconstantes"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadastro(9))
        rsSalvar.Fields(1) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(3) = X
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    
    SqlSalvar = "select * from tbmateriais where tbmateriais.idprd = '" & txtCadastro(9) & "'"
    rsSalvaMat.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvaMat.EOF Then rsSalvaMat.AddNew
    rsSalvaMat.Fields(0) = Val(txtCadastro(9))
    rsSalvaMat.Fields(1) = txtCadastro(2)
    rsSalvaMat.Fields(2) = txtCadastro(3)
    rsSalvaMat.Fields(3) = txtCadastro(6)
    rsSalvaMat.Fields(4) = txtCadastro(7)
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
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DesbloqueiaControles
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X) = ""
    Next
    ListView1.ListItems.Clear
    txtCadastro(0).Text = Format(GeraCodigo, "000000")
    txtCadastro(0).Enabled = False
    For X = 1 To txtCadastro.Count - 1
        txtCadastro(X).Enabled = True
    Next
    cboCadastro = ""
    ListView1.Enabled = True
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadastro(9).Text = Format(rsMaterial.Fields(0), "000000")
    txtCadastro(0).Text = rsMaterial.Fields(1)
    txtCadastro(1).Text = rsMaterial.Fields(2)
    If Not IsNull(rsMaterial.Fields(4)) Then txtCadastro(8).Text = rsMaterial.Fields(4)
    If Not IsNull(rsMaterial.Fields(5)) Then txtCadastro(2).Text = rsMaterial.Fields(5)
    If Not IsNull(rsMaterial.Fields(6)) Then txtCadastro(6).Text = rsMaterial.Fields(6)
    If Not IsNull(rsMaterial.Fields(7)) Then txtCadastro(3).Text = rsMaterial.Fields(7)
    If Not IsNull(rsMaterial.Fields(8)) Then txtCadastro(7).Text = rsMaterial.Fields(8)
    'cbocadastro.Text = rsMaterial(4)
    'BloqueiaControles
End Sub

Private Function ValidaGravacao()
    ValidaGravacao = False
    If txtCadastro(2).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(2).Tag, vbInformation, "Atenção"
        Me.txtCadastro(2).SetFocus
        Exit Function
    End If
    If txtCadastro(6).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(6).Tag, vbInformation, "Atenção"
        Me.txtCadastro(6).SetFocus
        Exit Function
    End If
    If txtCadastro(3).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(3).Tag, vbInformation, "Atenção"
        Me.txtCadastro(3).SetFocus
        Exit Function
    End If
    ValidaGravacao = True
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadastro(0).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(0).Tag, vbInformation, "Atenção"
        Me.txtCadastro(X).SetFocus
        Exit Function
    End If
    If txtCadastro(1).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(1).Tag, vbInformation, "Atenção"
        Me.txtCadastro(X).SetFocus
        Exit Function
    End If
    If txtCadastro(2).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(2).Tag, vbInformation, "Atenção"
        Me.txtCadastro(X).SetFocus
        Exit Function
    End If
    If txtCadastro(6).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtCadastro(6).Tag, vbInformation, "Atenção"
        Me.txtCadastro(X).SetFocus
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
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X).Enabled = False
    Next
    cboCadastro.Enabled = False
    ListView1.Enabled = False
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadastro.Count - 1
        txtCadastro(X).Enabled = True
    Next
    'cbocadastro.Enabled = True
    ListView1.Enabled = True
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirMaterial
    SqlGera = "Select top 1 * from tbMateriais order by codmaterial Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsMaterial.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharMaterial
End Function

Private Sub AbrirMaterial()
    SqlM = "Select * from tbMateriais Order by codmaterial"
    rsMaterial.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharMaterial()
    rsMaterial.Close
    Set rsMaterial = Nothing
End Sub

Private Sub ResultPesq()
    sqlMaterial = "select a.idprd,a.CODIGOPRD,a.NOMEFANTASIA,a.codtb2fat,c.descricao,b.formula,b.forpint,b.constpint,b.observacao from " & vBancoTotvs & ".dbo.TPRD as a left join " & sDatabaseName & ".dbo.tbMateriais as b on a.IDPRD = b.IDPRD left join " & vBancoTotvs & ".dbo.ttb2 as c on a.CODTB2FAT = c.CODTB2FAT where a.IDPRD = '" & Val(varGlobal) & "'"
    rsMaterial.Open sqlMaterial, cnBanco, adOpenKeyset, adLockReadOnly
    If rsMaterial.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Código do Material não encontrado", Ok, critico, "Atenção"
    End If
    rsMaterial.Close
    Set rsMaterial = Nothing
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
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = txtCadastro(2).Text
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = txtCadastro(6).Text
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
        txtCadastro(4).SetFocus
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
    Dim X As Integer, Y As Integer
    If ValidaCampo = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.Text1 Then
                Me.Text1 = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = Format(txtCadastro(4), "###,##0.000000000;(###,##0.000000000)")
                ListView1.SelectedItem.ListSubItems.Item(2) = Me.txtCadastro(5).Text
                Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
                
                txtCadastro(4) = ""
                txtCadastro(5) = ""
                Me.Text1 = ""
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , "const0(" & ListView1.ListItems.Count + 1 & ")")
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , "const0(" & ListView1.ListItems.Count + 1 & ")")
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = Format(txtCadastro(4), "###,##0.000000000;(###,##0.000000000)")
    ItemLst.SubItems(2) = Me.txtCadastro(5).Text
    Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
    txtCadastro(4) = ""
    txtCadastro(5) = ""
    Me.Text1 = ""
End Sub

Private Sub ExcluirItem()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    ListView1.ListItems.Remove (X)
End Sub

Private Sub AlterarItem()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.Text1.Text = ListView1.ListItems.Item(X)
    Me.txtCadastro(4).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.txtCadastro(5).Text = ListView1.SelectedItem.ListSubItems.Item(2)
End Sub

Private Sub ListView1_DblClick()
    AlterarItem
End Sub

