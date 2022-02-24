VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Métodos & Processos"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13845
   Icon            =   "frmMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular"
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
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Frame Frame8 
      Caption         =   "Observação "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   23
      Top             =   2280
      Width           =   7695
      Begin VB.TextBox txtformula 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1455
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Resultado "
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
      Left            =   240
      TabIndex        =   20
      Top             =   7200
      Width           =   7695
      Begin VB.TextBox txtResultado 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Decoder "
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
      Left            =   240
      TabIndex        =   19
      Top             =   6360
      Width           =   7695
      Begin VB.TextBox txtDecoder 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Variáveis "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   4200
      Width           =   7695
      Begin VB.TextBox txtformula 
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
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Tag             =   "Insira as variáveis de acordo com a Observação acima"
         ToolTipText     =   "Insira as variáveis de acordo com a Observação acima"
         Top             =   360
         Width           =   7455
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   3255
      Left            =   8040
      TabIndex        =   14
      Top             =   5760
      Width           =   5655
      Begin MSComctlLib.ListView ListView2 
         Height          =   2895
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5106
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   16777215
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
      Caption         =   "Referências "
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
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox txtformula 
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
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtformula 
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
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label Label4 
         Caption         =   "Fórmula:"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Parâmetros:"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fórmulas "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   8040
      TabIndex        =   8
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   16
         Top             =   4080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7223
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   120
      Picture         =   "frmMP.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   13
      Left            =   720
      Picture         =   "frmMP.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Centro de Custo "
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
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.TextBox txtformula 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   3
         Top             =   480
         Width           =   10575
      End
      Begin VB.TextBox txtformula 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
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
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Nome:"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "ID:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   7320
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "frmMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variaveis que irao receber os valores referente aos parametros das formulas
'para localizar os dados na tabela de classificação

'Variaveis que irão receber os dados da tabela de classificação após a localizacao
Private vTMedio As Double '
Private vFFadiga As Double
Private vOrganiza As Double
Private vSomaTempo As Double

'Variáveis que irão receber os dados do textBox de parametro para realizar a localização na
'tabela de parametros
Private vGrupo As String
Private vDimTipo As String
Private vDimValor As String
Private vInterTipo As String
Private vInterValor As String
Private vNomeA As String
Private vNomeB As String
Private vNomeC As String

Private var(50) As Double
Private cons(50) As Double

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        'Chama CC - Centro de Custo
        ChamaGrid "tbCCusto", "nome", txtformula(0), frmFormulaCC, "idprd", "nome"
        CarregaTxt "tbCCusto", "idprd", "S", "", "", txtformula(0), txtformula(1), 0, 1, txtformula(0), "S", txtformula(1)
        montaEstrutTreeview
        compoeDadosLVs
    Case 13 'Sair do formulário
        Unload Me
    End Select
End Sub

'Private Sub Command1_Click()
'    preparaDados
'End Sub

Private Sub preparaDados()
    LimpaVariaveis
    If txtformula(5) = "" Then
        MsgBox "Favor informar o campo: " & txtformula(5).Tag, vbInformation, "Atenção"
        txtformula(5).SetFocus
        Exit Sub
    End If
    'Calcula as formulas carregadas a partir das funções abaixo carregadas
    'a partir dos dados informados no campo de variáveis
    If Mid$(txtformula(2).Text, 1, 7) <> "formula" Then
        separaDadosVar txtformula(5)
        separaDadosPar txtformula(2)
        separaDadosCons
        localizaClassificacao
        substituiValores txtformula(3)
    Else
        If txtformula(7) <> "" Then
            'Acha o resultado referente a formula1
            separaDadosVar txtformula(5)
            separaDadosCons
            substituiValores txtformula(7)
            calculaValores 2
            localizaClassificacao
        End If
        
        If txtformula(10) <> "" Then
            'Acha o resultado referente a formula3
            separaDadosVar txtformula(5)
            separaDadosCons
            substituiValores txtformula(10)
            calculaValores 2
            localizaClassificacao
        End If
        
        vTMedio = Format(vSomaTempo, "#,##0.00;(#,##0.00)")
        'Pega o resultado das formulas 1 e 2 e aplica na formula3
        separaDadosVar txtformula(5)
        separaDadosPar txtformula(2)
        separaDadosCons
'        localizaClassificacao
        substituiValores txtformula(3)
    End If
End Sub

Private Sub Command2_Click()
    preparaDados
    txtResultado = ""
    calculaValores 1
End Sub

Private Sub Form_Load()
    listview_cabecalho
End Sub

Private Sub ListView1_Click()
    desmarcaCHK
    LimpaVariaveis
    LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
    compoeControles
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ListView1.ListItems(Item.Index).Selected = True
End Sub

Private Sub TreeView1_Click()
    AlteraTreeview
    LimpaVariaveis
    LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
    compoeDadosLVs
    compoeControles
End Sub

Private Sub txtformula_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            If txtformula(0).Text = "" Then
                MsgBox "Selecione primeiro um CC - Centro de Custo"
                Exit Sub
            End If
            CarregaTxt "tbCCusto", "idprd", "S", "", "", txtformula(0), txtformula(1), 0, 1, txtformula(0), "S", txtformula(1)
            montaEstrutTreeview
            compoeDadosLVs
        End If
    End Select
End Sub

Private Sub listview_cabecalho()

'    ListView1.ColumnHeaders.Clear
'    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 6
'    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 1.5
'    ListView1.ColumnHeaders.Add , , "Parametros", ListView1.Width / 10000
'
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Valor constante", ListView2.Width / 2.5
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 3.5
'
'    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub compoeDadosLVs()
    LimpaVariaveis
    LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
    'Faz referências a Funções que estão no: Module1.bas
'    'Listview1 - Formulas
'    LimpaLV ListView1
'    chamaSQL "select a.idform,a.nmform,a.parametros from tbFormula as a where a.idprd = '" & txtformula(0) & "'"
'    Compoe_Listview ListView1, Sqlp, "000"
    
'    ExcluirHipoLV ListView1
    
    'Listview2 - Constantes
    LimpaLV ListView2
    chamaSQL "Select a.idseq,a.valconst,a.descricao from tbconstantes as a where a.idprd = '" & txtformula(0) & "'"
    Compoe_Listview ListView2, Sqlp, "000"
End Sub

Private Sub ExcluirHipoLV(LV As ListView)
On Error Resume Next
    Dim X As Integer, Y As Integer
    Y = LV.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        If LV.SelectedItem.ListSubItems.Item(2) = "NV" Then
            LV.ListItems.Remove (X)
            Y = LV.ListItems.Count
            X = 1
        End If
    Next
End Sub

Private Sub compoeControles()
    Dim rsCompoe As New ADODB.Recordset
    Dim SqlCompoe As String
    'SqlCompoe = "Select a.parametros,a.formula from tbFormula as a where a.idprd = '" & txtformula(0) & "' and idform = '" & Val(txtformula(4)) & "'"
    SqlCompoe = "Select a.parametros,a.formula,a.observacao from tbFormula as a inner join tbproduto as b on a.idprd = b.idprd where a.idprd = '" & txtformula(0) & "' and idform = '" & Val(txtformula(4)) & "'"
    rsCompoe.Open SqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
        txtformula(2).Text = rsCompoe.Fields(0) 'Parâmetros
        txtformula(3).Text = rsCompoe.Fields(1) 'Formula
        If Not IsNull(rsCompoe.Fields(2)) Then txtformula(6).Text = rsCompoe.Fields(2) 'Observação
    Else
        txtformula(2).Text = "" 'Parâmetros
        txtformula(3).Text = "" 'Formula
        txtformula(6).Text = "" 'Observação
    End If
    If Mid$(txtformula(2).Text, 1, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 9, 1), 1
    End If
    If Mid$(txtformula(2).Text, 12, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 20, 1), 2
    End If
    rsCompoe.Close
    Set rsCompoe = Nothing
End Sub

Private Sub desmarcaCHK()
    Dim X As Integer, Y As Integer, J As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    J = ListView1.SelectedItem.Index
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Checked = True Then
            ListView1.ListItems.Item(X).Checked = False
        End If
    Next
    ListView1.ListItems.Item(J).Checked = True
    If ListView1.ListItems.Item(J).Checked = True Then ListView1.Enabled = True
    txtformula(4) = ListView1.ListItems.Item(J)
End Sub

'As 3 próximas SUBs são referentes a montagem e manipulação do TREEVIEW1
Private Sub montaEstrutTreeview()
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim vNome1 As String, vNome2 As String, vNome3 As String
    Dim nd As Node
    Dim vPula As Integer
    Dim vNo As Integer, vNo2 As Integer
    Dim vNomeNo As String
       
    TreeView1.Nodes.Clear

    SqlTreeview = "Select * from tbFormula as a where a.idprd = '" & txtformula(0) & "' order by a.idprd,a.nmform"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
    separaDadosTree rsTreeview.Fields(2)
    vNome1 = vNomeA
    vNome2 = vNomeB
    vNome3 = vNomeC
    vNo = 0
    On Error Resume Next
    Do While Not rsTreeview.EOF
        'PRIMEIRO NO
        Set nd = TreeView1.Nodes.Add(, , vNome1, vNome1)
        Do While vNome1 = vNomeA And Not rsTreeview.EOF
            If vNomeB <> "" Then
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
                'SEGUNDO NO
                Set nd = TreeView1.Nodes.Add(vNome1, tvwChild, vNomeNo, vNome2)
                Do While vNome2 = vNomeB And vNomeC <> "" And Not rsTreeview.EOF
                    'TERCEIRO NO
                    'OBS: OS VALORES DOS NOs NÃO PODEM SE REPETIR
                    'FOI ADICIONADO UM CONTADOR AO IDENTIFICADOR DO NO PARA QUE ELE NÃO SE REPITA
                    Set nd = TreeView1.Nodes.Add(vNomeNo, tvwChild, vNomeC & vNo, vNomeC)
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                    separaDadosTree rsTreeview.Fields(2)
                    vPula = 1
                Loop
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
            End If
            If vPula = 0 Then
                If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                separaDadosTree rsTreeview.Fields(2)
            End If
            vPula = 0
            
            If Not rsTreeview.EOF Then
                vNome2 = vNomeB
            End If
        Loop
        If Not rsTreeview.EOF Then
            vNome1 = vNomeA
            vNome2 = vNomeB
            vNome3 = vNomeC
        End If
    Loop
    rsTreeview.Close
    Set rsTreeview = Nothing
End Sub

Private Sub separaDadosTree(vTxtForm As String)
    Dim RECEBE As String
    Dim CONTADOR As Integer, X As Integer
    CONTADOR = 0
    vNomeA = ""
    vNomeB = ""
    vNomeC = ""
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If CONTADOR = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If CONTADOR = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
End Sub

Private Sub AlteraTreeview()
    Dim rsAlteraTreeview As New ADODB.Recordset
    Dim SqlAlteraTreeview As String
    
    
    Dim llng_Contador As Long
    Dim vNmNo As String
    
    
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            vNmNo = TreeView1.Nodes(llng_Contador).FullPath
        End If
    Next
    vNmNo = Replace(vNmNo, "\", ";")
    
    SqlAlteraTreeview = "Select idform,nmform,formula,parametros from tbFormula where nmform = '" & vNmNo & "'"
    rsAlteraTreeview.Open SqlAlteraTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsAlteraTreeview.EOF Then txtformula(4) = rsAlteraTreeview.Fields(0)

End Sub

'A função abaixo pega os valores dos parâmetro informados no textBox e armazena em variáveis
'específicas para cada valor
Private Sub separaDadosPar(vTxtForm As TextBox)
    Dim RECEBE As String
    Dim CONTADOR As Integer, vNum As Integer
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If CONTADOR = 0 And RECEBE <> "-" Then vGrupo = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 1 Then vDimTipo = RECEBE 'Variável vDimTipo receber o valor do segundo parâmetro
            If CONTADOR = 2 Then vDimValor = RECEBE 'Variavel vDimTipo recebe o valor do terceiro parâmetro
            If CONTADOR = 3 Then vInterTipo = RECEBE 'Variável vInterTipo recebe o valor do quarto parâmetro
            If CONTADOR = 4 Then vInterValor = RECEBE 'Variável vInterValor recebe o valor do quinto parâmetro
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If CONTADOR = 0 And RECEBE <> "-" Then vGrupo = RECEBE
    If CONTADOR = 1 Then vDimTipo = RECEBE
    If CONTADOR = 2 Then vDimValor = RECEBE
    If CONTADOR = 3 Then vInterTipo = RECEBE
    If CONTADOR = 4 Then vInterValor = RECEBE
    
    If Mid$(vDimValor, 1, 3) = "var" Then
        vNum = Val(Mid$(vDimValor, 5, 2))
        vDimValor = var(Val(Mid$(vDimValor, 5, 2)))
        vDimValor = Replace(vDimValor, ",", ".")
    End If
    If Mid$(vInterValor, 1, 3) = "var" Then
        vNum = Val(Mid$(vDimValor, 5, 2))
        vInterValor = var(Val(Mid$(vInterValor, 5, 2)))
        vInterValor = Replace(vInterValor, ",", ".")
    End If
End Sub

'A função abaixo pega os valores das variáveis informados no textBox txtformula(5) e armazena em Arrays: var(?)
'específicas para cada valor
Private Sub separaDadosVar(vTxtForm As TextBox)
    Dim RECEBE As String
    Dim CONTADOR As Integer, X As Integer
    CONTADOR = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If CONTADOR = 0 Then var(1) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 1 Then var(2) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 2 Then var(3) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 3 Then var(4) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 4 Then var(5) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 5 Then var(6) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If CONTADOR = 0 Then var(1) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 1 Then var(2) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 2 Then var(3) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 3 Then var(4) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 4 Then var(5) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 5 Then var(6) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
End Sub

'A função abaixo pega os valores das constantes informados no Listview2 e armazena em Arrays: cons(?)
'específicas para cada valor
Private Sub separaDadosCons()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True
        If ListView2.ListItems.Item(X).Selected = True Then
            cons(Val(ListView2.ListItems.Item(X))) = ListView2.SelectedItem.ListSubItems.Item(1)
        End If
    Next
End Sub

'Localiza a classificação na tabela baseado nos dados capturados na função separaDados
Private Sub localizaClassificacao()
    Dim rsLocaliza As New ADODB.Recordset
    Dim SqlLocaliza As String
    If vInterValor <> "" Then
        SqlLocaliza = "select * from tbClassificacao where idprd = '" & txtformula(0) & "' and idgrupo = '" & Val(vGrupo) & "' and '" & vDimValor & "' BETWEEN dim1 and dim2 AND '" & vInterValor & "' BETWEEN inter1 and inter2"
    End If
    If vInterValor = "" And vDimValor <> "" Then
        SqlLocaliza = "select * from tbClassificacao where idprd = '" & txtformula(0) & "' and idgrupo = '" & Val(vGrupo) & "' and '" & vDimValor & "' BETWEEN dim1 and dim2"
    End If
    
    If SqlLocaliza <> "" Then
        rsLocaliza.Open SqlLocaliza, cnBanco, adOpenKeyset, adLockReadOnly
        If Not rsLocaliza.EOF Then
            vTMedio = rsLocaliza.Fields(7)
            vFFadiga = rsLocaliza.Fields(8)
            vOrganiza = rsLocaliza.Fields(9)
            vSomaTempo = vSomaTempo + (var(2) / vTMedio)
            rsLocaliza.Close
            Set rsLocaliza = Nothing
        End If
    End If
End Sub

Private Sub substituiValores(vFormula As TextBox)
    Dim X As Integer
    Dim vPreserva As String
    vPreserva = ""
    vPreserva = vFormula
    For X = 1 To 50
        vFormula = Replace(vFormula, "cons(" & (X) & ")", cons(X))
        vFormula = Replace(vFormula, "var(" & (X) & ")", var(X))
        vFormula = Replace(vFormula, "vTMedio", vTMedio)
        vFormula = Replace(vFormula, "vFFadiga", vFFadiga)
        vFormula = Replace(vFormula, "vOrganiza", vOrganiza)
    Next
    vFormula = Replace(vFormula, ",", ".")
    txtDecoder = vFormula
    vFormula = vPreserva
End Sub

Private Sub calculaValores(vQual As Integer)
    'O ScriptControl é um componente. Ele interpreta e executa a formula/expressão numérica de um textbox
    If vQual = 1 Then
        txtResultado = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
    Else
        vGrupo = "1"
        vDimValor = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
        vDimValor = Replace(vDimValor, ",", ".")
        vDimValor = Replace(vDimValor, "(", "")
        vDimValor = Replace(vDimValor, ")", "")
        'MsgBox vResultFormula
    End If
End Sub

Private Sub LimpaVariaveis()
    vGrupo = ""
    vDimTipo = ""
    vDimValor = ""
    vInterTipo = ""
    vInterValor = ""
    vSomaTempo = 0
    vTMedio = 0
    vFFadiga = 0
    vOrganiza = 0
    vSomaTempo = 0
End Sub

'vPosicao indica a posicao da formula
Private Sub localizaFormula(vNForm As Integer, vPosicao As Integer)
    Dim rsFormula As New ADODB.Recordset
    Dim SqlFormula As String
    SqlFormula = "select * from tbFormula as a where a.idprd = '" & txtformula(0) & "' and idform = '" & vNForm & "'"
    rsFormula.Open SqlFormula, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsFormula.EOF Then
        If vPosicao = 1 Then
            txtformula(7).Text = rsFormula.Fields(4) 'Formula 2
            txtformula(8).Text = rsFormula.Fields(3) 'Parametros 2
        ElseIf vPosicao = 2 Then
            txtformula(10).Text = rsFormula.Fields(4) 'Formula 2
            txtformula(9).Text = rsFormula.Fields(3) 'Parametros 2
        End If
    End If
    rsFormula.Close
    Set rsFormula = Nothing
End Sub
