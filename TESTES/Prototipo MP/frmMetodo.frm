VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMetodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PrototipoX"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13665
   Icon            =   "frmMetodo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Nos Checados"
      Height          =   495
      Left            =   6960
      TabIndex        =   52
      Top             =   7440
      Width           =   1815
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6855
      Left            =   6840
      TabIndex        =   51
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   12091
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4920
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   120
      Picture         =   "frmMetodo.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Substituir"
      Height          =   495
      Left            =   4440
      TabIndex        =   21
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   5775
      Begin VB.TextBox txtDecoder 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   4440
      TabIndex        =   22
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resultado"
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
      Left            =   120
      TabIndex        =   27
      Top             =   7200
      Width           =   5775
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   5295
      Left            =   2400
      TabIndex        =   26
      Top             =   1800
      Width           =   1935
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   9
         Left            =   480
         TabIndex        =   20
         Text            =   "6"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   8
         Left            =   480
         TabIndex        =   19
         Text            =   "2"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   18
         Text            =   "4"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   6
         Left            =   480
         TabIndex        =   17
         Text            =   "2"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   16
         Text            =   "1"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   15
         Text            =   "8"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   14
         Text            =   "6"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Text            =   "5"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Text            =   "4"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtConst 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Text            =   "2"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "8"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Variáveis "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   25
      Top             =   1800
      Width           =   2055
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   9
         Left            =   480
         TabIndex        =   10
         Text            =   "6"
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   8
         Left            =   480
         TabIndex        =   9
         Text            =   "2"
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   8
         Text            =   "3"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   6
         Left            =   480
         TabIndex        =   7
         Text            =   "4"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   6
         Text            =   "6"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   5
         Text            =   "5"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Text            =   "3"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Text            =   "2"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Text            =   "5"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Text            =   "3"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "8"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fórmula"
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
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtFormula 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Text            =   "((var(1)*var(2)*var(6))+const(3)-const(7))/const(1)+20"
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmMetodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private var(10) As Double
Private cons(10) As Double
Private vNomeA As String
Private vNomeB As String
Private vNomeC As String


Private Sub Command4_Click()
    Dim X As Integer, vContador As Integer, vQtdNos As Integer
    vContador = 0
    X = 0
    vQtdNos = TreeView1.Nodes.Count
    For X = 1 To vQtdNos
        If TreeView1.Nodes.Item(X).Checked = True Then
            vContador = vContador + 1
        End If
    Next
    MsgBox vContador
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Command2_Click()
    substituiValores txtFormula
End Sub

Private Sub Command1_Click()
    calculaValores
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub substituiValores(vFormula As TextBox)
    Dim X As Integer
    Dim vPreserva As String
    vPreserva = vFormula
    For X = 1 To 10
        vFormula = Replace(vFormula, "const(" & (X) & ")", txtConst(X - 1))
        vFormula = Replace(vFormula, "var(" & (X) & ")", txtVar(X - 1))
    Next
    vFormula = Replace(vFormula, ",", ".")
    txtDecoder = vFormula
    vFormula = vPreserva
End Sub

Private Sub calculaValores()
    'O ScriptControl é um componente. Ele interpreta e executa a formula/expressão numérica de um textbox
    txtResultado = Format(ScriptControl1.Eval(txtDecoder), "#,##0.000;(#,##0.000)")
End Sub

Private Sub Form_Load()
    montaEstrutTreeview
End Sub

Private Sub TreeView1_DblClick()
    AlteraTreeview
End Sub

Private Sub txtConst_GotFocus(Index As Integer)
    mudaCorText txtConst(Index)
End Sub

Private Sub txtConst_LostFocus(Index As Integer)
    voltaCorText txtConst(Index)
End Sub

Private Sub txtVar_GotFocus(Index As Integer)
    mudaCorText txtVar(Index)
End Sub

Private Sub txtVar_LostFocus(Index As Integer)
    voltaCorText txtVar(Index)
End Sub

Private Sub montaEstrutTreeview()
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim vNome1 As String, vNome2 As String, vNome3 As String
    Dim nd As Node
    Dim vPula As Integer
    Dim vNo As Integer, vNo2 As Integer
    Dim vNomeNo As String
       
    
    TreeView1.Nodes.Clear

    SqlTreeview = "Select * from tbFormula order by idprd,nmform"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    
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
End Sub

