VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro de movimentações"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "frmFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      TabIndex        =   22
      Top             =   3720
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   6360
      TabIndex        =   21
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7335
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5530
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Text            =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Novo Filtro"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Limite máximo de linhas"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   5880
      Width           =   2295
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin ZEUS.chameleonButton cmdFiltro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   6120
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
      MICON           =   "frmFiltro.frx":0CCA
      PICN            =   "frmFiltro.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdFiltro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   6120
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
      MICON           =   "frmFiltro.frx":19C0
      PICN            =   "frmFiltro.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Configurar colunas"
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
      TabIndex        =   2
      Top             =   5760
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro obrigatório "
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
      Top             =   4680
      Width           =   7335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmFiltro.frx":26B6
         Left            =   120
         List            =   "frmFiltro.frx":26B8
         TabIndex        =   1
         Tag             =   "Lista de opções do filtro"
         ToolTipText     =   "Lista de opções do filtro"
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select:"
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
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
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
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "where:"
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
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Group:"
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
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Order:"
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
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   615
      Left            =   1320
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   16095
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   16095
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   16095
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   16095
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   16095
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Declare Function GetScrollInfo Lib "user32" (ByVal HWnd As Long, ByVal fnBar As Long, lpScrollInfo As SCROLLINFO) As Long
 
'interestingly, API Viewer doesn't have these constants, translating from Windows.h is straight forward
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
  
'my declarations
Private Const c_EntryTxt = ""
Private m_ColIndex As Long 'listview col index
Private m_RowIndex As Long 'listview row index
'Acima - usado poder editar o listview --------------------

Private vPonte1 As TextBox, vPonte2 As TextBox, vPonte3 As TextBox, vPonte4 As TextBox, vPonte5 As TextBox, vPonte6 As TextBox
Private vQuery As String
Private vSubstituto As String
Private vContaActive As Integer

Private Sub cmdFiltro_Click(Index As Integer)
On Error GoTo Err
    Select Case Index
    Case 0
        If Combo1.Text = "" Then
            mobjMsg.Abrir "Nenhum filtro selecionado", Ok, critico
            Exit Sub
        End If
        If Check1.Value = 0 Then checaFiltro = False Else checaFiltro = True
        Tipo = True
        FiltroGeral = Combo1.Text
        gravaLog "Filtro - Tipo: " & Combo1.Text, "", ""
        LimiteLinhas = Val(Text1.Text)
        montaLV1
        Unload Me
        Set frmFiltro = Nothing
    Case 1
        Tipo = False
        Unload Me
        Set frmFiltro = Nothing
    End Select
    Exit Sub
Err:
    Resume Next
End Sub

Private Sub Command1_Click()
    ExcluirItemLV ListView2
    DeletaFiltro
    Combo1.Text = ""
End Sub

Private Sub Command2_Click()
    frmCriaFiltro.Show 1
    CarregaFiltro
End Sub

Private Sub Form_Activate()
    CarregaFiltro
End Sub

Private Sub Form_Load()
On Error Resume Next
    vponteiro = 1
    
'----------------------------------------
    
     vContaActive = vContaActive + 1
    If vContaActive = 1 Then
        Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
        Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
        Set vPonte3 = Me.Controls.Add("VB.TextBox", "vPonte3")
        Set vPonte4 = Me.Controls.Add("VB.TextBox", "vPonte4")
        Set vPonte5 = Me.Controls.Add("VB.TextBox", "vPonte5")
        Set vPonte6 = Me.Controls.Add("VB.TextBox", "vPonte6")
    End If
    
    If vFil = "N" Then
        Combo1.Enabled = False
    End If
    If dataFilter1 = "" Then
        'DTPicker1 = Format("01/01/" & Year(Date), "dd/mm/yyyy")
        'DTPicker2 = Format("31/12/" & Year(Date), "dd/mm/yyyy")
    Else
        'DTPicker1 = dataFilter1
        'DTPicker2 = dataFilter2
    End If
    Text1.Text = LimiteLinhas
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub CarregaFiltro()
On Error GoTo Err
    Dim rsListaQuerys As New ADODB.Recordset
    Dim SqlListaQuerys As String
    SqlListaQuerys = "select a.nomefiltro,a.query,a.expressao,a.tipofiltro,a.usuario,a.modulo,a.padrao from tbfiltro as a where a.usuario = '" & NomUsu & "' and a.modulo = '" & Formulario & "' or a.modulo = '" & Formulario & "' and a.tipofiltro = 'global'"
                     '"select a.nomefiltro,a.query,a.expressao,a.tipofiltro,a.usuario,a.modulo,a.padrao from tbfiltro where a.usuario '" & NomUsu & "' and a.modulo = '" & Formulario & "' or a.modulo = '" & Formulario & "' and a.tipofiltro = 'global'"
                     '"Select a.nomefiltro,a.query,a.expressao,a.tipofiltro,a.usuario,a.modulo,a.padrao from tbFiltro as a where a.usuario = '" & NomUsu & "' and a.modulo = '" & Formulario & "'"
    rsListaQuerys.Open SqlListaQuerys, cnBanco, adOpenKeyset, adLockReadOnly
    listview_cabecalho
    If rsListaQuerys.RecordCount = 0 Then
        vPonte1 = SqlLV
        vPonte2 = "Todos"
        vPonte3 = "global"
        vPonte4 = NomUsu
        vPonte5 = Formulario
        vPonte6 = "S"
        ListView2.View = lvwReport 'Modo de Exibição do seu Listview
        IncluirLV ListView2, vPonte2, vPonte1, Text2, vPonte3, vPonte4, vPonte5, vPonte6, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2
        montaLV1
        gravaFiltro
    Else
        Combo1.Clear
        ListView2.ListItems.Clear
        chamaSQL SqlListaQuerys
        Compoe_Listview ListView2, Sqlp, "00"
        Unload chamaForm
        montaLV1
    End If
    rsListaQuerys.Close
    Set rsListaQuerys = Nothing
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

Private Sub Form_Unload(Cancel As Integer)
    vContaActive = 0
End Sub

Private Sub ListView2_Click()
On Error Resume Next
    vPonte1.Text = Combo1.Text
    AlteraLV ListView2, vPonte1, vPonte2, Text3, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2
    Combo1.Text = vPonte1.Text
    ListView2.ListItems.Item(vponteiro).Selected = True
    ListView2.SetFocus
End Sub

Private Sub ListView2_DblClick()
    If Combo1.Text = "" Then
        mobjMsg.Abrir "Nenhum filtro selecionado", Ok, critico
        Exit Sub
    End If
    If Check1.Value = 0 Then checaFiltro = False Else checaFiltro = True
    Tipo = True
    FiltroGeral = Combo1.Text
    gravaLog "Filtro - Tipo: " & Combo1.Text, "", ""
    LimiteLinhas = Val(Text1.Text)
    montaLV1
    Unload Me
    Set frmFiltro = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        If Not IsNumeric(Chr(KeyAscii)) Then
            If Check1.Value = 0 Then checaFiltro = False Else checaFiltro = True
            Tipo = True
            FiltroGeral = Combo1.Text
            gravaLog "Filtro - Tipo: " & Combo1.Text, "", ""
            LimiteLinhas = Val(Text1.Text)
            Unload Me
            Set frmFiltro = Nothing
        End If
    End If
End Sub

'----------------------------------------
'----------------------------------------
'----------------------------------------
'----------------------------------------

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 1.3
    ListView2.ColumnHeaders.Add , , "Query", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "Expressao", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "Tipo Filtro", ListView2.Width / 5
    
    ListView2.ColumnHeaders.Add , , "usuario", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "modulo", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "padrao", ListView2.Width / 10000
    
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub montaLV1()
    vQuery = ""
    Executar vQuery 'Text1.Text
    vSubstituto = vNovoFiltro
    LocalString SqlLV
    
    frmFiltro.Label6.Caption = Replace(frmFiltro.Label6.Caption, "top 500", "top " & LimiteLinhas)
    
    If Label9.Caption = "Label9" Then
        If frmFiltro.Label6.Caption <> "Label6" And vSubstituto <> " " Then SqlLV = frmFiltro.Label6.Caption & " " & frmFiltro.Label7 & " where " & vSubstituto
    Else
        If frmFiltro.Label6.Caption <> "Label6" And vSubstituto <> " " Then SqlLV = frmFiltro.Label6.Caption & " " & frmFiltro.Label7 & " where " & vSubstituto & " " & Label9
    End If
    SeparaDados
End Sub

Private Sub SeparaDados()
    On Error GoTo Err
    Dim vPoints(4, 1) As String
    Dim RECEBE As String
    Dim Contador As Integer, K As Integer
    K = 0
    vPoints(0, 0) = "from"
    vPoints(0, 1) = "4"
    vPoints(1, 0) = "where"
    vPoints(1, 1) = "5"
    vPoints(2, 0) = "group"
    vPoints(2, 1) = "5"
    vPoints(3, 0) = "order"
    vPoints(3, 1) = "5"
'    MsgBox "Você digitou:" & Len(Text1) & " caracteres"
    Contador = 0
    For X = 1 To Len(SqlLV)
        If Mid(SqlLV, X, vPoints(K, 1)) = vPoints(K, 0) Then
            If K = 0 Then Label6 = RECEBE
            If K = 1 Then Label7 = RECEBE
            If K = 2 Then Label8 = RECEBE
            If K = 3 Then Label9 = RECEBE
            If K = 4 Then Label10 = RECEBE
            K = K + 1
            RECEBE = ""
            X = X - 1
        Else
            RECEBE = RECEBE & Mid(SqlLV, X, 1)
        End If
    Next
    If K = 0 Then Label6 = RECEBE
    If K = 1 Then Label7 = RECEBE
    If K = 2 Then Label8 = RECEBE
    If K = 3 Then Label9 = RECEBE
    If K = 4 Then Label10 = RECEBE
    Exit Sub
Err:
    Exit Sub
End Sub

Private Function Executar(vSql As String)
On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Function
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    
    Dim rsExecutaSel As New ADODB.Recordset
    Dim SqlExecutaSel As String
    SqlExecutaSel = "Select a.query,a.expressao from tbFiltro as a where a.usuario = '" & ListView2.SelectedItem.ListSubItems.Item(4) & "' and a.modulo = '" & ListView2.SelectedItem.ListSubItems.Item(5) & "' and a.nomefiltro = '" & Combo1.Text & "'"
    rsExecutaSel.Open SqlExecutaSel, cnBanco, adOpenKeyset, adLockReadOnly
    If rsExecutaSel.RecordCount > 0 Then
        SqlLV = rsExecutaSel.Fields(0)
        vNovoFiltro = rsExecutaSel.Fields(1)
    Else
        SqlLV = ListView2.SelectedItem.ListSubItems.Item(1)
        vNovoFiltro = ListView2.SelectedItem.ListSubItems.Item(2)
    End If
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

Private Sub LocalString(vQuery As String)
    Dim vContador As Integer
    
    If InStr(UCase(vNovoFiltro), UCase("like '[]'")) > 0 Then
        vContador = 1
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("like '[]'")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, UCase("like '[]"), UCase("like ") & "'%" & vAlteraLike)
                vNovoFiltro = vSubstituto
                vContador = vContador + 1
            Else
                vSubstituto = Replace(vNovoFiltro, UCase("like '[]"), UCase("like ") & "'%" & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
    End If
    
    If InStr(UCase(vNovoFiltro), UCase("like '[datetime]'")) > 0 Then
        vContador = 1
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("like '[datetime]'")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "LIKE '[datetime]", UCase("='") & vAlteraLike)
                vNovoFiltro = vSubstituto
                vContador = vContador + 1
            Else
                vSubstituto = Replace(vNovoFiltro, "LIKE '[datetime]", UCase("='") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
'----------------
    End If
    If InStr(UCase(vNovoFiltro), UCase("BETWEEN")) > 0 Then
        vContador = 1
        frmPassaParametro.Text2.Visible = True
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("'[datetime")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "[datetime1]", UCase("") & vAlteraLike)
                vContador = vContador + 1
            Else
                vNovoFiltro = vSubstituto
                vSubstituto = Replace(vNovoFiltro, "[datetime2]", UCase("") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
    End If
    If InStr(UCase(vNovoFiltro), UCase("'[datetime")) > 0 Then
        vContador = 1
        frmPassaParametro.Text2.Visible = True
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("'[datetime")) > 0
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "'[datetime1]", UCase("'") & vAlteraLike)
                vContador = vContador + 1
            Else
                vNovoFiltro = vSubstituto
                vSubstituto = Replace(vNovoFiltro, "'[datetime2]", UCase("'") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
'-----------------
    End If
    If InStr(UCase(vNovoFiltro), UCase("IN([])")) > 0 Then
        vContador = 1
        frmPassaParametro.Show 1
        While InStr(UCase(vSubstituto), UCase("IN([])")) > 0
            vAlteraLike = Replace(vAlteraLike, "%", "")
            If vContador = 1 Then
                vSubstituto = Replace(vNovoFiltro, "[]", UCase("") & vAlteraLike)
                vNovoFiltro = vSubstituto
                vContador = vContador + 1
            Else
                vNovoFiltro = vSubstituto
                vSubstituto = Replace(vNovoFiltro, "[]", UCase("") & vAlteraLike2)
                vNovoFiltro = vSubstituto
            End If
        Wend
    End If
End Sub

Private Sub gravaFiltro()
On Error GoTo Err
    Dim rsGravaFiltro As New ADODB.Recordset
    Dim SqlGravaFiltro As String

    SqlGravaFiltro = "Select * from tbFiltro "
    rsGravaFiltro.Open SqlGravaFiltro, cnBanco, adOpenKeyset, adLockOptimistic
    rsGravaFiltro.AddNew
    rsGravaFiltro.Fields(1) = vPonte4
    rsGravaFiltro.Fields(2) = vPonte5
    rsGravaFiltro.Fields(3) = vPonte3
    rsGravaFiltro.Fields(4) = vPonte2
    rsGravaFiltro.Fields(5) = vPonte1
    rsGravaFiltro.Fields(6) = ""
    rsGravaFiltro.Fields(7) = vPonte6
    rsGravaFiltro.Update
    rsGravaFiltro.Close
    Set rsGravaFiltro = Nothing
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

Private Sub DeletaFiltro()
On Error GoTo Err
    Dim rsDeletarFiltro As New ADODB.Recordset
    Dim SqlDeletarFiltro As String
    SqlDeletarFiltro = "Delete from tbFiltro where modulo = '" & Formulario & "' and nomefiltro = '" & Combo1.Text & "'"
    rsDeletarFiltro.Open SqlDeletarFiltro, cnBanco
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

'--------------------------------------
' EDITA LISTVIEW
'----------------------------------------------------
'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------
Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, leftPos As Single 'the left pos of the column
    Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate
    'ENTRA ABAIXO SOMENTE SE ESTIVER NO MODULO DE SERVICOS
    If Button = vbRightButton Then 'vbLeftButton Then
        If Not ListView2.SelectedItem Is Nothing Then
            ListView2.LabelEdit = lvwManual
            dx = GetLvwDeltaX
            lvwX = X + dx
            'Função da coluna que altera o status do requesito (possui/não possui)
            For i = 4 To 4
                leftPos = ListView2.Left + ListView2.ColumnHeaders(i).Left
                If lvwX > leftPos And lvwX < leftPos + ListView2.ColumnHeaders(i).Width Then 'we found the column
                    m_RowIndex = ListView2.SelectedItem.Index 'row
                    m_ColIndex = i 'column
                        AlteraLVFiltro i
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub txtLvw_LostFocus()
'On Error GoTo TrataErro
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView2.ListItems(m_RowIndex).Text = Trim(txtLvw.Value) 'put in the text
    ElseIf m_ColIndex = 7 Then
        If txtLvw.Value = Date Then
            ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = ""
        Else
            ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Value)
        End If
    ElseIf m_ColIndex Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Value)
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    ListView2.SetFocus
TrataErro:
    Exit Sub
End Sub

Function GetLvwDeltaX() As Single
    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView2.ColumnHeaders(ListView2.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView2.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView2.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
    Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
    Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
    
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
        Set lvwCol = ListView1.ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView2.Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView2.Left Then txtLeft = ListView2.Left + 48
    
        txtRightMax = ListView2.Left + ListView2.Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView2.ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView2.Left + ListView2.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView2.Top
        If Not ListView2.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView2.Top + ListView2.Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView2.Top + ListView2.SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
        With txtLvw '.move produces runtime error with -ve values
            .Left = txtLeft
            .Top = txtTop '+ 4450
            .Width = txtWidth - 100
            .Height = ListView2.SelectedItem.Height - 100
        End With
    End If
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub AlteraLVFiltro(coluna As Integer)
'    On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            If ListView2.SelectedItem.ListSubItems.Item(3) = "individual" Then
                ListView2.SelectedItem.ListSubItems.Item(3) = "global"
            Else
                ListView2.SelectedItem.ListSubItems.Item(3) = "individual"
            End If
            Exit For
        End If

    Next
    varGlobal = ListView2.ListItems.Item(X)
    
    Dim rsAlteraTipo As New ADODB.Recordset
    Dim SqlAlteraTipo As String
    SqlAlteraTipo = "update tbfiltro set tipofiltro = '" & ListView2.SelectedItem.ListSubItems.Item(3) & "' where nomefiltro = '" & ListView2.ListItems.Item(X) & "' and usuario = '" & ListView2.SelectedItem.ListSubItems.Item(4) & "' and modulo = '" & ListView2.SelectedItem.ListSubItems.Item(5) & "'"
    rsAlteraTipo.Open SqlAlteraTipo, cnBanco
    
    Exit Sub
Err:
    varGlobal = ""
    mobjMsg.Abrir "Erro", Ok, critico, "SAF"
    Exit Sub
End Sub

