VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADPModelo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modelo ADP"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmADPModelo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.ComboBox txtLvw 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmADPModelo.frx":0CCA
      Left            =   6960
      List            =   "frmADPModelo.frx":0CD7
      TabIndex        =   11
      Text            =   "Institucional"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Marque os itens que serão avaliados nesse treinamento"
      Height          =   6735
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   6135
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   11033
         View            =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modelos de avaliação"
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtModelo 
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtModelo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtModelo 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "Nome do modelo de avaliação"
         ToolTipText     =   "Nome do modelo de avaliação"
         Top             =   1080
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmADPModelo.frx":0D01
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "frmADPModelo.frx":0D69
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmADPModelo.frx":0DD5
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7858
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MAESTRO.chameleonButton cmdModelo 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Excluir modelo"
         ToolTipText     =   "Excluir modelo"
         Top             =   1440
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
         MICON           =   "frmADPModelo.frx":0E41
         PICN            =   "frmADPModelo.frx":0E5D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdModelo 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Tag             =   "Editar modelo"
         ToolTipText     =   "Editar modelo"
         Top             =   1440
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
         MICON           =   "frmADPModelo.frx":1B37
         PICN            =   "frmADPModelo.frx":1B53
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdModelo 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   7
         Tag             =   "Novo modelo"
         ToolTipText     =   "Novo modelo"
         Top             =   1440
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
         MICON           =   "frmADPModelo.frx":282D
         PICN            =   "frmADPModelo.frx":2849
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdModelo 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Tag             =   "Incluir modelo"
         ToolTipText     =   "Incluir modelo"
         Top             =   1440
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
         MICON           =   "frmADPModelo.frx":3523
         PICN            =   "frmADPModelo.frx":353F
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
   Begin MAESTRO.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   1
      Left            =   1320
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
      MICON           =   "frmADPModelo.frx":4219
      PICN            =   "frmADPModelo.frx":4235
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
      Index           =   2
      Left            =   720
      TabIndex        =   17
      Tag             =   "Carregar MODELO"
      ToolTipText     =   "Carregar MODELO"
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
      MICON           =   "frmADPModelo.frx":4F0F
      PICN            =   "frmADPModelo.frx":4F2B
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
      MICON           =   "frmADPModelo.frx":5C05
      PICN            =   "frmADPModelo.frx":5C21
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
Attribute VB_Name = "frmADPModelo"
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

Private Sub cmdModelo_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirModelo
        LimpaControlesModelo
        desmarcaTodos
    Case 1
        LimpaControlesModelo
        desmarcaTodos
    Case 2
        editaControlesModelo
        desmarcaTodos
    End Select
End Sub

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        GravarDados
        vCodModeloAval = Val(txtModelo(2))
    Case 1
        Unload Me
        Set frmAvaliacaoProg = Nothing
    Case 2
        carregarModelo
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    listview_cabecalho
    txtModelo(2) = Format(vCodModeloAval, "00")
    compoeLV2
    marcaLV2
    LimpaControlesModelo
    CompoeAD
    'CompoePontosAE
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Avaliação", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Dimensão", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Peso", ListView1.Width / 10000
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 1.5
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub compoeLV2()
    Dim rsModelo As New ADODB.Recordset
    Dim sqlModelo As String
    sqlModelo = "select * from tbModeloADP where codcoligada = '" & vCodcoligada & "'"
    rsModelo.Open sqlModelo, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView2.ListItems.Clear
    While Not rsModelo.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsModelo.Fields(0), "000")) 'codigo do modelo de avaliação
        ItemLst.SubItems(1) = "" & rsModelo.Fields(1) 'nome do modelo de avaliação
        rsModelo.MoveNext
        X = X + 1
    Wend
    rsModelo.Close
    Set rsModelo = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub marcaLV2()
    If txtModelo(3) = "" Then Exit Sub
    Dim X As Integer, Y As Integer, J As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If Val(ListView2.ListItems.Item(X)) = Val(txtModelo(3)) Then
            ListView2.ListItems.Item(X).Checked = True
            Exit Sub
        End If
    Next
End Sub

Private Sub CompoeAD()
    Dim rsAE As New ADODB.Recordset
    Dim sqlAE As String
    'Tabela que armazena os itens que serão avaliados na Avaliação de Eficácia do Treinamento
    sqlAE = "select * from tbAvaliacao where codcoligada = '" & vCodcoligada & "' and tipo = 'AD' and ativo = 'S'"
    rsAE.Open sqlAE, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView1.ListItems.Clear
    
    'Essa partte da rotina é p eu definir a altura da linha do listview
    '------------------------------------------------------------------
    ImageList1.ImageHeight = 20 ' rode com este valor (altura da linha)
    ImageList1.ListImages.Add , , Me.Icon
    Set ListView1.SmallIcons = ImageList1
    '------------------------------------------------------------------
   
    While Not rsAE.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsAE.Fields(0), "000")) 'codigo da avaliação AE
        ItemLst.SubItems(1) = "" & rsAE.Fields(1) 'nome da avaliação AE
        ItemLst.SubItems(2) = "-" 'Dimensão
        ItemLst.SubItems(3) = "" & rsAE.Fields(5) 'Dimensão
        ItemLst.SubItems(4) = "" & rsAE.Fields(3) 'Peso
        rsAE.MoveNext
        X = X + 1
    Wend
    rsAE.Close
    Set rsAE = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE PONTUAÇÃO DA AVALIAÇÃO DE EFICÁCIA <<<<<<<<<<
Private Sub CompoePontosAE()
    Dim rsHab As New ADODB.Recordset
    Dim sqlHab As String
    sqlHab = "Select * from tbModeloADPItens where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & Val(txtModelo(3)) & "'"
    rsHab.Open sqlHab, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If rsHab.RecordCount = 0 Then Exit Sub
    Y = ListView1.ListItems.Count
    While Not rsHab.EOF
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True
            If Val(ListView1.ListItems.Item(X)) = rsHab.Fields(0) Then
                ListView1.ListItems.Item(X).Checked = True
                ListView1.SelectedItem.ListSubItems.Item(2) = rsHab.Fields(2)
            End If
        Next
        rsHab.MoveNext
    Wend
    rsHab.Close
    Set rsHab = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub carregarModelo()
    If chamaForm.Caption = "Sistema" Then
        frmConfSistema.SkinLabel18 = txtModelo(0)
        If ListView2.ListItems.Count > 0 Then
            For X = 1 To ListView2.ListItems.Count
                If ListView2.ListItems.Item(X).Checked = True Then
                    ListView2.ListItems.Item(X).Selected = True
                    frmConfSistema.SkinLabel18 = Val(ListView2.ListItems.Item(X))
                End If
            Next
        End If
    Else
        Dim Y As Integer
        Y = ListView1.ListItems.Count
        chamaForm.ListView1.ListItems.Clear
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Checked = True Then
                ListView1.ListItems.Item(X).Selected = True
                Dim ItemLst As ListItem
                Dim K As Integer, L As Integer, LV3Edit As String
                L = ListView1.ListItems.Count
                If L > 0 Then
                    Set ItemLst = chamaForm.ListView1.ListItems.Add(, , Format(ListView1.ListItems.Item(X), "000000"))
                    ItemLst.SubItems(1) = ListView1.SelectedItem.ListSubItems.Item(1) ' Nome avaliação
                    ItemLst.SubItems(2) = ListView1.SelectedItem.ListSubItems.Item(4) ' Peso
                    ItemLst.SubItems(3) = "0"  ' Nota
                    ItemLst.SubItems(4) = ListView1.SelectedItem.ListSubItems.Item(2) ' Dimensão
                    ItemLst.SubItems(5) = ListView1.SelectedItem.ListSubItems.Item(3) ' Descrição
                    ItemLst.ListSubItems(3).Bold = True
                End If
            End If
        Next
    End If
    mobjMsg.Abrir "Modelo foi ativo com sucesso", Ok, informacao, "SGC"
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    Dim Y As Integer, X As Integer, K As Integer
    cnBanco.BeginTrans
    '>>>>>> GRAVAR AVALIACAO <<<<<<<<<
    sqlDeletar = "Delete from tbModeloADP where codmodelo = '" & Val(txtModelo(3)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbModeloADP"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    If ListView2.ListItems.Count > 0 Then
        For X = 1 To ListView2.ListItems.Count
            If ListView2.ListItems.Item(X).Checked = True Then
                ListView2.ListItems.Item(X).Selected = True
                rsSalvar.AddNew
                rsSalvar.Fields(0) = Val(ListView2.ListItems.Item(X))
                rsSalvar.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
                rsSalvar.Fields(2) = vCodcoligada 'Codigo da coligada
            End If
        Next
    End If
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    sqlDeletar = "Delete from tbModeloADPItens where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & Val(txtModelo(3)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbModeloADPItens where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & Val(txtModelo(3)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    If ListView1.ListItems.Count > 0 Then
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.ListItems.Item(X).Checked = True Then
                rsSalvar.Find "codavaliacao=" & "'" & Val(ListView1.ListItems.Item(X)) & "'"
                If rsSalvar.EOF Then rsSalvar.AddNew
                rsSalvar.Fields(0) = Val(ListView1.ListItems.Item(X))
                rsSalvar.Fields(1) = Val(txtModelo(3))
                rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
                rsSalvar.Fields(3) = vCodcoligada 'Codigo da coligada
            End If
        Next
    End If
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    cnBanco.CommitTrans
    mobjMsg.Abrir "Os dados foram salvos com sucesso", Ok, informacao, "SGC"
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
TrataErro1:
    Resume Next
End Sub

'-------- ROTINAS DE MODELOS DE AVALIAÇÃO DE EFICÁCIA -------------
Private Sub IncluirModelo()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaModelo = False Then Exit Sub
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView2.ListItems.Item(X) = Me.txtModelo(0) Then
                Me.txtModelo(0) = ListView2.ListItems.Item(X)
                ListView2.SelectedItem.ListSubItems.Item(1) = txtModelo(1)
                Y = ListView2.ListItems.Count
                Me.ListView2.SortOrder = lvwAscending
                Exit Sub
            End If
        Next
        Set ItemLst = ListView2.ListItems.Add(, , txtModelo(0))
        Y = ListView2.ListItems.Count
    Else
        Set ItemLst = ListView2.ListItems.Add(, , txtModelo(0))
        Y = ListView2.ListItems.Count
        Me.ListView2.SortOrder = lvwAscending
    End If
    ItemLst.SubItems(1) = txtModelo(1)
    txtModelo(1).SetFocus
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Function ValidaModelo()
    ValidaModelo = False
    If txtModelo(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtModelo(1).Tag, Ok, critico, "Atenção"
        Me.txtModelo(1).SetFocus
        Exit Function
    End If
    ValidaModelo = True
End Function

Private Sub LimpaControlesModelo()
    Dim X As Integer
    txtModelo(1) = ""
    If ListView2.ListItems.Count > 0 Then
        txtModelo(0).Text = Format(GeraCodigo1, "000")
    Else
        txtModelo(0).Text = Format(Val(txtModelo(0)) + 1, "000")
    End If
End Sub

Private Function GeraCodigo1()
    Dim X As Integer
    X = 1
    Me.ListView2.SortOrder = lvwDescending
    ListView2.ListItems.Item(X).Selected = True
    GeraCodigo1 = Val(ListView2.ListItems.Item(X)) + 1
    Me.ListView2.SortOrder = lvwAscending
    Exit Function
End Function

Private Sub ListView2_Click()
   'pegaClick
   desmarcaModelos
   CompoeAD
   CompoePontosAE
End Sub

Private Sub desmarcaModelos()
    Dim X As Integer, Y As Integer, J As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    J = ListView2.SelectedItem.Index
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Item(X).Checked = False
        End If
    Next
    ListView2.ListItems.Item(J).Checked = True
    If ListView2.ListItems.Item(J).Checked = True Then ListView1.Enabled = True
    txtModelo(3) = ListView2.ListItems.Item(J)
End Sub

Private Sub desmarcaTodos()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        ListView2.ListItems.Item(X).Checked = False
    Next
    ListView1.Enabled = False
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ListView2.ListItems(Item.Index).Selected = True
End Sub

Private Sub ListView2_DblClick()
    editaControlesModelo
End Sub

Private Sub editaControlesModelo()
    If ListView2.ListItems.Count = 0 Then Exit Sub
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtModelo(0).Text = ListView2.ListItems.Item(X)
    Me.txtModelo(1).Text = ListView2.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub excluirControlesModelo()
    Dim rsVSePodeExcluir As New ADODB.Recordset
    Dim SqlVSePodeExcluir As String
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    Dim llng_Contador As Long
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    SqlVSePodeExcluir = "Select codmodelo from tbprogramacao where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & ListView2.ListItems.Item(X) & "'"
    rsVSePodeExcluir.Open SqlVSePodeExcluir, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVSePodeExcluir.RecordCount > 0 Then
        mobjMsg.Abrir "Existem programações para esse modelo. Não pode ser excluido", Ok, informacao, "Atenção"
        Exit Sub
    End If
    ListView2.ListItems.Remove (X)
End Sub

'**********************************************
'**********************************************
'**********************************************
'**********************************************
'**********************************************

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 3 To 3
            leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView1.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView1.SelectedItem.Text
                    Else
                        .Text = ListView1.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    If txtLvw.Enabled = True Then .SetFocus Else txtModelo(1).SetFocus
                End With
                Exit For
            End If
        Next i
    End If
End If
End Sub

Function GetLvwDeltaX() As Single
    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView1.ColumnHeaders(ListView1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView1.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
    Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
    Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
    
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
        Set lvwCol = ListView1.ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView1.Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView1.Left Then txtLeft = ListView1.Left + 48
    
        txtRightMax = ListView1.Left + ListView1.Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView1.ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView1.Left + ListView1.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView1.Top
        If Not ListView1.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView1.Top + ListView1.Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView1.Top + ListView1.SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
        With txtLvw '.move produces runtime error with -ve values
            .Left = 8225
            .Top = txtTop + 115
            .Width = 1250
            '.Height = 315
        End With
    End If
End Sub

Private Sub txtLvw_GotFocus()
    If txtLvw.Text = "" Then txtLvw.Text = " "
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView1 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
        'other keys can be used for navigation
    End Select
    If txtLvw.Text = "-" Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw <> "" And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLvw_LostFocus()
On Error GoTo TrataErro
    'AKI - desenvolver rotina para verificar qtd digitada
    If txtLvw.Text = " " Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw.Text <> "" And Len(txtLvw) = 1 Then txtLvw.Text = "-"
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView1.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView1.ListItems(ListView1.ListItems.Count) <> c_EntryTxt Then ListView1.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    'If ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) = "-" And ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) < ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) Then
    '    ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
    '    Exit Sub
    'End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    'If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > Val(ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2)) Then
    '    ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
    'End If
    
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    'txtModelo(1).SetFocus
    'ListView1.SetFocus
TrataErro:
    Exit Sub
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub txtModelo_Change(Index As Integer)
    Select Case Index
    Case 1
        CompoeAD
        desmarcaTodos
    End Select
End Sub

Private Sub txtModelo_Click(Index As Integer)
    Select Case Index
    Case 1
        CompoeAD
        desmarcaTodos
    End Select
End Sub

Private Sub txtModelo_GotFocus(Index As Integer)
    Select Case Index
    Case 1
        CompoeAD
        desmarcaTodos
    End Select
End Sub
