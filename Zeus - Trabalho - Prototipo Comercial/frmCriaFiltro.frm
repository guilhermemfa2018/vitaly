VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmCriaFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCriaFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Expressão "
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
      TabIndex        =   17
      Top             =   6480
      Width           =   7455
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "frmCriaFiltro.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         Picture         =   "frmCriaFiltro.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         Picture         =   "frmCriaFiltro.frx":265E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tipo do Filtro "
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
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option3 
            Caption         =   "Individual"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Global"
            Height          =   255
            Left            =   1560
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4048
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
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
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "frmCriaFiltro.frx":3328
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10080
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nome do FIltro: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Frame Frame6 
         Caption         =   "Tabela"
         Height          =   735
         Left            =   2040
         TabIndex        =   14
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox Text3 
            Height          =   330
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "E/OU"
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
         TabIndex        =   11
         Top             =   4440
         Width           =   1695
         Begin VB.OptionButton Option2 
            Caption         =   "OU"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "E"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   4335
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Visible         =   0   'False
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operador"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   4335
         Begin VB.ComboBox Combo1 
            Height          =   345
            ItemData        =   "frmCriaFiltro.frx":3FF2
            Left            =   240
            List            =   "frmCriaFiltro.frx":401D
            TabIndex        =   7
            Text            =   "="
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Campo"
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
         TabIndex        =   4
         Top             =   840
         Width           =   4335
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tipo"
         Height          =   735
         Left            =   2040
         TabIndex        =   1
         Top             =   5400
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox Text6 
            Height          =   330
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   11033
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
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
Attribute VB_Name = "frmCriaFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPonte1 As TextBox, vPonte2 As TextBox, vPonte3 As TextBox, vPonte4 As TextBox, vPonte5 As TextBox, vPonte6 As TextBox

Private Sub Combo1_Click()
    If Combo1.Text = "BETWEEN" Then
        Text5.Visible = True
    Else
        Text5.Visible = False
    End If
End Sub

Private Sub Command1_Click()
    
    If Mid$(vTabela1, 1, 9) <> "CORPORERM" Then
        If UCase(vTabela1) = UCase(Text3.Text) Then Text3 = "a."
    Else
        If UCase(Mid$(vTabela1, 15, 50)) = UCase(Text3.Text) Then Text3 = "a."
    End If
    If Mid$(vTabela2, 1, 9) <> "CORPORERM" Then
        If UCase(vTabela2) = UCase(Text3.Text) Then Text3 = "b."
    Else
        If UCase(Mid$(vTabela2, 15, 50)) = UCase(Text3.Text) Then Text3 = "B."
    End If
    If Mid$(vTabela3, 1, 9) <> "CORPORERM" Then
        If UCase(vTabela3) = UCase(Text3.Text) Then Text3 = "c."
    Else
        If UCase(Mid$(vTabela3, 15, 50)) = UCase(Text3.Text) Then Text3 = "c."
    End If
    If Mid$(vTabela4, 1, 9) <> "CORPORERM" Then
        If UCase(vTabela4) = UCase(Text3.Text) Then Text3 = "d."
    Else
        If UCase(Mid$(vTabela4, 15, 50)) = UCase(Text3.Text) Then Text3 = "d."
    End If
    If Mid$(vTabela5, 1, 9) <> "CORPORERM" Then
        If UCase(vTabela5) = UCase(Text3.Text) Then Text3 = "e."
    Else
        If UCase(Mid$(vTabela5, 15, 50)) = UCase(Text3.Text) Then Text3 = "e."
    End If
    
    
    
    If Text6.Text = "datetime" And Combo1.Text = "LIKE" Or Text6.Text = "datetime" And Combo1.Text = "NOT LIKE" Then
        Text2.Text = Mid$(Text2.Text, 1, 1) & Text6.Text & Mid$(Text2.Text, 2, 1)
    ElseIf Text6.Text = "datetime" And Combo1.Text = "BETWEEN" Then
        Text2.Text = Mid$(Text2.Text, 1, 1) & Text6.Text & "1" & Mid$(Text2.Text, 2, 1)
        Text5.Text = Mid$(Text5.Text, 1, 1) & Text6.Text & "2" & Mid$(Text5.Text, 2, 1)
    ElseIf Text6.Text = "datetime" And Text2.Text = "[]" Then
        Dim vConta As Integer
        vConta = ListView2.ListItems.Count + 1
        Text2.Text = Mid$(Text2.Text, 1, 1) & Text6.Text & vConta & Mid$(Text2.Text, 2, 1)
    ElseIf Text6.Text = "varchar" And Combo1.Text = "LIKE" Or Text6.Text = "varchar" And Combo1.Text = "NOT LIKE" Then
        If Text2.Text <> "[]" Then Text2.Text = "%" & Text2.Text & "%"
    ElseIf Combo1.Text = "IN" Or Combo1.Text = "NOT IN" Then
        Text2.Text = "(" & Text2.Text & ")"
    ElseIf Combo1.Text = "IS NULL" Or Combo1.Text = "IS NOT NULL" Then
        Text2.Text = Text2.Text
    End If
    
    
    If Text5.Visible = False Then
        If Combo1.Text <> "IN" And Combo1.Text <> "NOT IN" And Combo1.Text <> "IS NULL" And Combo1.Text <> "IS NOT NULL" Then
            vPonte1 = Text3.Text & Text1.Text & " " & Combo1.Text & " '" & Text2.Text & "'"
        Else
            vPonte1 = Text3.Text & Text1.Text & " " & Combo1.Text & Text2.Text
        End If
    Else
        vPonte1 = Text3.Text & Text1.Text & " " & Combo1.Text & " '" & Text2.Text & "' and '" & Text5.Text & "'"
    End If
        
    If ListView2.ListItems.Count = 0 Then
        vPonte2 = ""
    Else
        If Option1.Value = True Then
            vPonte2 = "AND"
        Else
            vPonte2 = "OR"
        End If
    End If
    IncluirLV ListView2, vPonte2, vPonte1, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2
    LimpaControles Text2, Text5, Text2, Text2, Text2, Text2, Text2, Text2, Text2, Text2
End Sub

Private Sub Command3_Click()
    ExcluirItemLV ListView2
End Sub

Private Sub Command4_Click()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    vNovoFiltro = ""
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        vNovoFiltro = vNovoFiltro & " " & ListView2.ListItems.Item(X) & " " & ListView2.SelectedItem.ListSubItems.Item(1)
    Next
    'frmFiltro.Text2 = frmFiltro.Label6.Caption & " " & frmFiltro.Label7 & " where " & vNovoFiltro
    If frmFiltro.Label9 <> "Label9" Then
        vPonte2 = frmFiltro.Label6.Caption & " " & frmFiltro.Label7 & " where " & vNovoFiltro & " " & frmFiltro.Label9
    Else
        vPonte2 = frmFiltro.Label6.Caption & " " & frmFiltro.Label7 & " where " & vNovoFiltro
    End If
    
    vPonte1 = vNovoFiltro
    If Option3.Value = True Then
        vPonte3 = "individual"
    Else
        vPonte3 = "global"
    End If
    vPonte4 = NomUsu
    vPonte5 = Formulario
    vPonte6 = "N"
    gravaFiltro
    'IncluirLV frmFiltro.ListView2, Text4, vPonte2, vPonte1, vPonte3, vPonte4, vPonte5, vPonte6, Text4, Text4, Text4, Text4, Text4, Text4, Text4, Text4
    frmFiltro.ListView2.ListItems.Clear
    Unload Me
End Sub

Private Sub Form_Load()
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    Set vPonte3 = Me.Controls.Add("VB.TextBox", "vPonte3")
    Set vPonte4 = Me.Controls.Add("VB.TextBox", "vPonte4")
    Set vPonte5 = Me.Controls.Add("VB.TextBox", "vPonte5")
    Set vPonte6 = Me.Controls.Add("VB.TextBox", "vPonte6")

    listview_cabecalho
    chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns " & _
             "Where TABLE_NAME = '" & vTabela1 & "' or " & _
             "TABLE_NAME = '" & vTabela2 & "' or " & _
             "TABLE_NAME = '" & vTabela3 & "' or " & _
             "TABLE_NAME = '" & vTabela4 & "' or " & _
             "TABLE_NAME = '" & vTabela5 & "' or " & _
             "TABLE_NAME = '" & vTabela6 & "' or " & _
             "TABLE_NAME = '" & vTabela7 & "' or " & _
             "TABLE_NAME = '" & vTabela8 & "' or " & _
             "TABLE_NAME = '" & vTabela9 & "' or " & _
             "TABLE_NAME = '" & vTabela10 & "'"
    Compoe_Listview ListView1, Sqlp, "00"
    tabelaRM
    
    Unload chamaForm
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub tabelaRM()
On Error GoTo Err
    ConexaoTotvs
    If Mid$(vTabela1, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela1, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela2, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela2, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela3, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela3, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela4, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela4, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela5, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela5, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela6, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela6, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela7, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela7, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela8, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela8, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela9, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela9, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    If Mid$(vTabela10, 1, 9) = "CORPORERM" Then
        chamaSQL "SELECT COLUMN_NAME, DATA_TYPE, '',TABLE_NAME From INFORMATION_SCHEMA.Columns Where TABLE_NAME = '" & Mid$(vTabela10, 15, 50) & "'"
        Compoe_Listview ListView1, Sqlp, "TOTVS"
    End If
    cnBancoTotvs.Close
    Set cnBancoTotvs = Nothing
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

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Campo", ListView1.Width / 1.3
    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 100000
    ListView1.ColumnHeaders.Add , , "Tamanho", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Tabela", ListView1.Width / 10000
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Operador", ListView2.Width / 7
    ListView2.ColumnHeaders.Add , , "Expressão", ListView2.Width / 2
    ListView2.ColumnHeaders.Add , , "Tabela", ListView2.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ListView1_DblClick()
    vPonte1.Text = Combo1.Text
    AlteraLV ListView1, Text1, Text6, Text3, Text3, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1
    'Combo1.Text = vPonte1.Text
End Sub

Private Sub Text2_LostFocus()
    If IsDate(Text2) Then
        Text2 = Format(Text2, "yyyy-mm-dd")
    End If
End Sub

Private Sub Text5_LostFocus()
    If IsDate(Text5) Then
        Text5 = Format(Text5, "yyyy-mm-dd")
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
    rsGravaFiltro.Fields(4) = Text4.Text
    rsGravaFiltro.Fields(5) = vPonte2
    rsGravaFiltro.Fields(6) = vPonte1
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
