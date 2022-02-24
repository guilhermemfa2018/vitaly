VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPesqger2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmPesqger2.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin IMRM.chameleonButton chameleonButton3 
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   240
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
      MICON           =   "frmPesqger2.frx":0CCA
      PICN            =   "frmPesqger2.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin IMRM.chameleonButton chameleonButton2 
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   5280
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
      MICON           =   "frmPesqger2.frx":19C0
      PICN            =   "frmPesqger2.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin IMRM.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5280
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
      MICON           =   "frmPesqger2.frx":26B6
      PICN            =   "frmPesqger2.frx":26D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa"
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
      TabIndex        =   6
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmPesqger2.frx":33AC
         Left            =   120
         List            =   "frmPesqger2.frx":33AE
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
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
Attribute VB_Name = "frmPesqger2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsTabela As New ADODB.Recordset
Dim sql As String

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    If procnom = "chamaCD" Then
        ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 5
        ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 1.4
    End If
    
    If apontaLV = 0 Or apontaLV = 12 Then
        ListView1.ColumnHeaders.Add , , "Chapa", ListView1.Width / 5
        ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 1.4
    Else
        ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 5
        ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 1.4
    End If
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub CompoeLVPesquisa()
    Dim ItemLst As ListItem
    Dim sql As String
    sql = Sqlp
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear
    While Not rsTabela.EOF
        If procnom = "chamaCD" Then
            Set ItemLst = ListView1.ListItems.Add(, , Format(rsTabela.Fields(Campo1), "000000")) 'Codigo
        Else
            Set ItemLst = ListView1.ListItems.Add(, , rsTabela.Fields(Campo1)) 'Codigo
        End If
        ItemLst.SubItems(1) = rsTabela.Fields(campo) 'Descricao
        rsTabela.MoveNext
    Wend
    rsTabela.Close
    'Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    Set rsTabela = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 1
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub chameleonButton1_Click()
    capturaDados
End Sub

Private Sub chameleonButton2_Click()
    Unload Me
End Sub

Private Sub chameleonButton3_Click()
    Pesquisar
End Sub

Private Sub Form_Activate()
    Me.Text1.SetFocus
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_Load()
    listview_cabecalho
    CompoeLVPesquisa
    CompoeComboLVPesq Combo1, ListView1, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'rsTabela.Close
    'Set rsTabela = Nothing
End Sub

'As duas Subs abaixo faz com que ordene o listview pela coluna que vc clicar
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView1, ColumnHeader
    Combo1.Text = ColumnHeader.Text
End Sub

Public Sub ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
    With ListView1
    If .SortKey <> Column.Index - 1 Then
        .SortKey = Column.Index - 1
        .SortOrder = lvwAscending
    Else
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End If
    .Sorted = -1
    End With
End Sub

Private Sub ListView1_DblClick()
    capturaDados
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        capturaDados
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        Pesquisar ' Sub que realiza a Pesquisa no Listview mediante ao que foi digitado no TexBox Text1 e ao q foi selecionado no ComboBox Combo1
    End If
End Sub

Private Sub Pesquisar(Optional Column As ColumnHeader = Nothing)
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count 'Conta as linhas preenchidas do Listview
    If Y > 0 Then 'Entra nessa condição se o Listview não estiver vazio
        Dim c As ColumnHeader
        Dim numCol As Integer
        numCol = 0
        For Each c In ListView1.ColumnHeaders
            If Combo1.Text = c Then Exit For
            numCol = numCol + 1
        Next
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True 'Seleciona a linha de acordo com o valor de "X"
            'SE FOR SELECIONADO A PRIMEIRA COLUNA
            If Combo1.Text = "" Then
                'Se não for selecionado nada no ComboBox Combo1
                Msgbox "Nenhum filtro de pesquisa selecionado"
                Exit Sub
            End If
            If numCol = 0 Then
                If UCase(ListView1.ListItems.Item(X)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    ListView1.SetFocus
                    Exit Sub
                End If
            'SE FOR SELECIONADO A PARTIR DA SEGUNDA COLUNA
            ElseIf numCol > 0 Then
                If UCase(ListView1.SelectedItem.ListSubItems.Item(numCol)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    ListView1.SetFocus
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub capturaDados()
On Error Resume Next
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Pesquisa = ListView1.ListItems.Item(X)
    Unload Me
End Sub

