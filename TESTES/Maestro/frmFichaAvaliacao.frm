VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Begin VB.Form frmFichaAvaliacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Avaliação"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmFichaAvaliacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7800
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
      MICON           =   "frmFichaAvaliacao.frx":0CCA
      PICN            =   "frmFichaAvaliacao.frx":0CE6
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
      TabIndex        =   6
      Tag             =   "Confirmar"
      ToolTipText     =   "Confirmar"
      Top             =   7800
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
      MICON           =   "frmFichaAvaliacao.frx":19C0
      PICN            =   "frmFichaAvaliacao.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtLvw 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colaborador/candidato"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin ACTIVESKINLibCtl.SkinLabel Text5 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmFichaAvaliacao.frx":26B6
         TabIndex        =   19
         Top             =   960
         Width           =   6975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFichaAvaliacao.frx":271C
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text2 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmFichaAvaliacao.frx":278A
         TabIndex        =   17
         Top             =   480
         Width           =   6975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFichaAvaliacao.frx":27F0
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmFichaAvaliacao.frx":2854
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFichaAvaliacao.frx":28BC
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmFichaAvaliacao.frx":292C
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFichaAvaliacao.frx":2994
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   1440
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Avaliação do treinamento"
      TabPicture(0)   =   "frmFichaAvaliacao.frx":29FA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Eficácia do treinamento"
      TabPicture(1)   =   "frmFichaAvaliacao.frx":2A16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Observação"
      TabPicture(2)   =   "frmFichaAvaliacao.frx":2A32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAvaliacao"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtAvaliacao 
         Height          =   5535
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   480
         Width           =   8655
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   5535
         Index           =   1
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9763
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
      Begin MSComctlLib.ListView ListView 
         Height          =   5535
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9763
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
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "00,00"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "00,00"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Pontos AE:"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Pontos AT:"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   7800
      Width           =   975
   End
End
Attribute VB_Name = "frmFichaAvaliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Abaixo - usado poder editar o listview --------------------
'straight from the standard API Viewver
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
'-------------------------------------------------------------

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        GravarDados
        VerificaDados
        Unload Me
    Case 1
        mobjMsg.Abrir "Deseja sair da ficha de avaliação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Pesquisa = 0
            Unload Me
            Set frmFichaAvaliacao = Nothing
        End If
    End Select
End Sub

Private Sub VerificaDados()
    Dim Y As Integer
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = ListView(0).ListItems.Count
    vqtdava = 0
    For X = 1 To Y
        ListView(0).ListItems(X).Selected = True
        If ListView(0).SelectedItem.ListSubItems.Item(3) > 0 Then vqtdava = vqtdava + 1
    Next
    If SSTab1.TabEnabled(1) = False Then
        If vqtdava = Y Then
            vsituacao = "Aprovado"
        Else
            vsituacao = "-"
        End If
        vNota = "-"
    Else
        If Val(Label6) < vAprovadoRest Then
            vsituacao = "Reprovado"
        ElseIf Val(Label6) >= aprovadorest And Val(Label6) < MediaGlobal Then
            vsituacao = "Aprovado com restrição"
        Else
            vsituacao = "Aprovado"
        End If
        vNota = Label6
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    If Legenda = "" Then SSTab1.TabEnabled(1) = False
    listview_cabecalho
    carregaDados
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
    ListView(0).ColumnHeaders.Clear
    ListView(1).ColumnHeaders.Clear
    ListView(0).ColumnHeaders.Add , , "Código", ListView(0).Width / 10
    ListView(0).ColumnHeaders.Add , , "Avaliação", ListView(0).Width / 3
    ListView(0).ColumnHeaders.Add , , "Peso", ListView(0).Width / 11
    ListView(0).ColumnHeaders.Add , , "Nota", ListView(0).Width / 11
    
    ListView(1).ColumnHeaders.Add , , "Código", ListView(1).Width / 10
    ListView(1).ColumnHeaders.Add , , "Avaliação", ListView(1).Width / 3
    ListView(1).ColumnHeaders.Add , , "Peso", ListView(1).Width / 11
    ListView(1).ColumnHeaders.Add , , "Nota", ListView(1).Width / 11
    
    ListView(0).View = lvwReport 'Modo de Exibição do seu Listview
    ListView(1).View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub carregaDados()
    Dim rsColab As New ADODB.Recordset
    Dim sqlColab As String
    sqlColab = "select a.cpf,a.nomecolaborador,a.codcolaborador,a.tipo,b.obsavaliacao from tbcolaboradores as a left join tbPendentesCur as b on a.cpf ='" & Mid(varGlobal2, 1, 11) & "' and b.cpf = '" & Mid(varGlobal2, 1, 11) & "' where a.codcoligada = '" & vCodcoligada & "' and b.codtreinamento = '" & Val(Mid$(varGlobal2, 18, 6)) & "' and b.ativo ='S'"
    rsColab.Open sqlColab, cnBanco, adOpenKeyset, adLockOptimistic
    Text1 = Mid(Mid(varGlobal2, 1, 11), 1, 3) + "." + Mid(Mid(varGlobal2, 1, 11), 4, 3) + "." + Mid(Mid(varGlobal2, 1, 11), 7, 3) + "-" + Mid(Mid(varGlobal2, 1, 11), 10, 2)
    Text2 = rsColab.Fields(1)
    Text4 = rsColab.Fields(2)
    Text5 = rsColab.Fields(3)
    If Not IsNull(rsColab.Fields(4)) Then txtAvaliacao = rsColab.Fields(4)
    rsColab.Close
    Set rsColab = Nothing
    Label5 = Label5 & "%"
    Label6 = Label6 & "%"
    CompoeAT
    CompoeAE
    CompoePontosATAE
    SomaLV 0
    SomaLV 1
End Sub

Private Sub CompoeAT()
    Dim rsAT As New ADODB.Recordset
    Dim SqlAT As String
    SqlAT = "select * from tbAvaliacao where codcoligada = '" & vCodcoligada & "' and tipo = 'AT' and ativo = 'S'"
    rsAT.Open SqlAT, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView(0).ListItems.Clear
    While Not rsAT.EOF
        Set ItemLst = ListView(0).ListItems.Add(, , Format(rsAT.Fields(0), "000")) 'codigo da avaliação AT
        ItemLst.SubItems(1) = "" & rsAT.Fields(1) 'nome da avaliação AT
        ItemLst.SubItems(2) = "" & rsAT.Fields(3) 'peso da avaliação AT
        ItemLst.SubItems(3) = "" & 0 'Nota da avaliação AT
        ItemLst.ListSubItems(3).Bold = True
        rsAT.MoveNext
        X = X + 1
    Wend
    rsAT.Close
    Set rsAT = Nothing
    Me.ListView(0).Sorted = True
    Me.ListView(0).SortKey = 0
    Me.ListView(0).SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE PONTUAÇÃO DO LISTVIEW(0) DA GUIA DE AVALIAÇÃO <<<<<<<<<<
Private Sub CompoePontosATAE()
    Dim rsAvaliacaoTrei As New ADODB.Recordset
    Dim sqlAvaliacaoTrei As String
    
    sqlAvaliacaoTrei = "Select a.*,b.tipo,b.peso from tbAvaliacaoTrei as a inner join tbAvaliacao as b on a.codcoligada = '" & vCodcoligada & "' and a.codavaliacao = b.codavaliacao where b.tipo = 'AT' and a.cpf = '" & Mid(varGlobal2, 1, 11) & "' and a.codprogramacao = '" & Val(Mid(varGlobal2, 12, 6)) & "'order by a.codavaliacao"
    rsAvaliacaoTrei.Open sqlAvaliacaoTrei, cnBanco, adOpenKeyset, adLockOptimistic
    Dim X As Integer, Y As Integer
    Y = ListView(0).ListItems.Count
    While Not rsAvaliacaoTrei.EOF
        For X = 1 To Y
            ListView(0).ListItems(X).Selected = True
            If Val(ListView(0).ListItems.Item(X)) = rsAvaliacaoTrei.Fields(2) Then
                ListView(0).SelectedItem.ListSubItems.Item(3) = rsAvaliacaoTrei.Fields(3)
            End If
        Next
        rsAvaliacaoTrei.MoveNext
    Wend
    rsAvaliacaoTrei.Close
    
    sqlAvaliacaoTrei = "Select a.*,b.tipo,b.peso from tbAvaliacaoTrei as a inner join tbAvaliacao as b on a.codcoligada = '" & vCodcoligada & "' and a.codavaliacao = b.codavaliacao where b.tipo = 'AE' and a.cpf = '" & Mid(varGlobal2, 1, 11) & "' and a.codprogramacao = '" & Val(Mid(varGlobal2, 12, 6)) & "'order by a.codavaliacao"
    rsAvaliacaoTrei.Open sqlAvaliacaoTrei, cnBanco, adOpenKeyset, adLockOptimistic
    Y = ListView(1).ListItems.Count
    While Not rsAvaliacaoTrei.EOF
        For X = 1 To Y
            ListView(1).ListItems(X).Selected = True
            If Val(ListView(1).ListItems.Item(X)) = rsAvaliacaoTrei.Fields(2) Then
                ListView(1).SelectedItem.ListSubItems.Item(3) = rsAvaliacaoTrei.Fields(3)
            End If
        Next
        rsAvaliacaoTrei.MoveNext
    Wend
    rsAvaliacaoTrei.Close
    Set rsAvaliacaoTrei = Nothing
    
    Me.ListView(0).Sorted = True
    Me.ListView(0).SortKey = 0
    Me.ListView(0).SortOrder = lvwAscending
    Me.ListView(1).Sorted = True
    Me.ListView(1).SortKey = 0
    Me.ListView(1).SortOrder = lvwAscending
End Sub

Private Sub CompoeAE()
    Dim rsAE As New ADODB.Recordset
    Dim sqlAE As String
    sqlAE = "select * from tbAvaliacao as a inner join tbAvaliacaoprog as b on a.codcoligada = '" & vCodcoligada & "' and a.codavaliacao = b.codavaliacao where tipo = 'AE' and ativo = 'S' and b.codmodelo = '" & Val(vCodModeloAval) & "'"
    rsAE.Open sqlAE, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView(1).ListItems.Clear
    While Not rsAE.EOF
        Set ItemLst = ListView(1).ListItems.Add(, , Format(rsAE.Fields(0), "000")) 'codigo da avaliação AE
        ItemLst.SubItems(1) = "" & rsAE.Fields(1) 'nome da avaliação AE
        ItemLst.SubItems(2) = "" & rsAE.Fields(3) 'peso da avaliação AE
        ItemLst.SubItems(3) = "" & 0 'Nota da avaliação AE
        ItemLst.ListSubItems(3).Bold = True
        rsAE.MoveNext
        X = X + 1
    Wend
    rsAE.Close
    Set rsAE = Nothing
    Me.ListView(1).Sorted = True
    Me.ListView(1).SortKey = 0
    Me.ListView(1).SortOrder = lvwAscending
End Sub

Private Sub SomaLV(ind As Integer)
    Dim X As Integer, Y As Integer, F As Integer
    Y = ListView(ind).ListItems.Count
    Dim somaPeso As Double, somaPontos As Double
    somaPeso = 0
    somaPontos = 0
    For X = 1 To Y
        If ListView(ind).ListItems.Item(X).Selected = True Then F = X
    Next
    For X = 1 To Y
        ListView(ind).ListItems.Item(X).Selected = True
        somaPeso = somaPeso + ListView(ind).SelectedItem.ListSubItems.Item(2)
        somaPontos = somaPontos + ListView(ind).SelectedItem.ListSubItems.Item(3)
    Next
    If somaPontos <> 0 Or somaPeso <> 0 Then
        If ind = 0 Then Label5 = Format(somaPontos * 100 / somaPeso, "#,##00.00;(#,##0.00)") & "%"
        If ind = 1 Then Label6 = Format(somaPontos * 100 / somaPeso, "#,##00.00;(#,##0.00)") & "%"
        ListView(ind).ListItems.Item(F).Selected = True
    End If
End Sub

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------
Private Sub ListView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, leftPos As Single 'the left pos of the column
    Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate
    If Button = vbLeftButton Then
        'Listview(0)
        If Not ListView(0).SelectedItem Is Nothing Then
            ListView(0).LabelEdit = lvwManual
            dx = GetLvwDeltaX(0)
            lvwX = X + dx
            For i = 4 To 4
                leftPos = ListView(0).Left + ListView(0).ColumnHeaders(i).Left
                If lvwX > leftPos And lvwX < leftPos + ListView(0).ColumnHeaders(i).Width Then 'we found the column
                    m_RowIndex = ListView(0).SelectedItem.Index 'row
                    m_ColIndex = i 'column
                    MoveTxtLvw dx 'move and size the edit box over the selected item
                    With txtLvw 'turn on edit box
                        If i = 1 Then 'copy the text of the selected item to txtlvw
                            .Text = ListView(0).SelectedItem.Text
                        Else
                            .Text = ListView(0).SelectedItem.SubItems(i - 1)
                        End If
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SetFocus
                    End With
                    Exit For
                End If
            Next i
            SomaLV 0
        End If
        'Listview(1)
        If Not ListView(1).SelectedItem Is Nothing Then
            ListView(1).LabelEdit = lvwManual
            dx = GetLvwDeltaX(1)
            lvwX = X + dx
            For i = 4 To 4
                leftPos = ListView(1).Left + ListView(1).ColumnHeaders(i).Left
                If lvwX > leftPos And lvwX < leftPos + ListView(1).ColumnHeaders(i).Width Then 'we found the column
                    m_RowIndex = ListView(1).SelectedItem.Index 'row
                    m_ColIndex = i 'column
                    MoveTxtLvw dx 'move and size the edit box over the selected item
                    With txtLvw 'turn on edit box
                        If i = 1 Then 'copy the text of the selected item to txtlvw
                            .Text = ListView(1).SelectedItem.Text
                        Else
                            .Text = ListView(1).SelectedItem.SubItems(i - 1)
                        End If
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SetFocus
                    End With
                    Exit For
                End If
            Next i
            SomaLV 1
        End If
    End If
End Sub

Function GetLvwDeltaX(ind As Integer) As Single
    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView(ind).ColumnHeaders(ListView(ind).ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView(ind).HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView(ind).Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
    Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
    Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
If SSTab1.Tab = 0 Then
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX(0) 'called from subclass event
        Set lvwCol = ListView(0).ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView(0).Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView(0).Left Then txtLeft = ListView(0).Left + 48
    
        txtRightMax = ListView(0).Left + ListView(0).Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView(0).ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView(0).Left + ListView(0).ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView(0).Top
        If Not ListView(0).HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView(0).Top + ListView(0).Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView(0).Top + ListView(0).SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
    
        With txtLvw '.move produces runtime error with -ve values
            If txtLeft < 11000 Then .Left = txtLeft + 205 Else .Left = txtLeft - 140
            .Top = txtTop + 1570
            .Width = txtWidth - 3500
            .Height = ListView(0).SelectedItem.Height - 8
        End With
    End If
ElseIf SSTab1.Tab = 1 Then
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX(1) 'called from subclass event
        Set lvwCol = ListView(1).ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView(1).Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView(1).Left Then txtLeft = ListView(1).Left + 48
    
        txtRightMax = ListView(1).Left + ListView(1).Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView(1).ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView(1).Left + ListView(1).ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView(1).Top
        If Not ListView(1).HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView(1).Top + ListView(1).Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView(1).Top + ListView(1).SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
    
        With txtLvw '.move produces runtime error with -ve values
            If txtLeft < 11000 Then .Left = txtLeft + 205 Else .Left = txtLeft - 140
            .Top = txtTop + 1570
            .Width = txtWidth - 3500
            .Height = ListView(1).SelectedItem.Height - 8
        End With
    End If
End If
End Sub

Private Sub txtLvw_GotFocus()
    If txtLvw.Text = "" Then txtLvw.Text = " "
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView(0) is edited
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
If SSTab1.Tab = 0 Then
    'AKI - desenvolver rotina para verificar qtd digitada
    If txtLvw.Text = " " Then txtLvw.Text = ""
    
    If Not IsNumeric(txtLvw.Text) And txtLvw.Text <> "" And Len(txtLvw) = 1 Then txtLvw.Text = "0"
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView(0).ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView(0).ListItems(ListView(0).ListItems.Count) <> c_EntryTxt Then ListView(0).ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView(0).ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > Val(ListView(0).ListItems(m_RowIndex).SubItems(m_ColIndex - 2)) Then
        ListView(0).ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
        mobjMsg.Abrir "Valor acima do peso", Ok, critico, "Atenção"
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    ListView(0).SetFocus
ElseIf SSTab1.Tab = 1 Then
    'AKI - desenvolver rotina para verificar qtd digitada
    If txtLvw.Text = " " Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw.Text <> "" And Len(txtLvw) = 1 Then txtLvw.Text = "0"
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView(1).ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView(0).ListItems(ListView(0).ListItems.Count) <> c_EntryTxt Then ListView(0).ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView(1).ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > Val(ListView(1).ListItems(m_RowIndex).SubItems(m_ColIndex - 2)) Then
        ListView(1).ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
        mobjMsg.Abrir "Valor acima do peso", Ok, critico, "Atenção"
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    ListView(1).SetFocus
End If
TrataErro:
    Exit Sub
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView(0)'s vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView(0).HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub GravarDados()
'On Error GoTo TrataErro
    'If ValidaCampo = False Then Exit Sub
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim rsSalvarObs As New ADODB.Recordset
    Dim SqlSalvarObs As String
    
    Dim Y As Integer, X As Integer, K As Integer
    cnBanco.BeginTrans
    '>>>>>> GRAVAR AVALIACAO <<<<<<<<<
    SqlSalvar = "Select * from tbavaliacaotrei where codcoligada = '" & vCodcoligada & "' and cpf = '" & Mid(varGlobal2, 1, 11) & "' and codprogramacao = '" & Val(Mid(varGlobal2, 12, 6)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    If ListView(0).ListItems.Count > 0 Then
        For X = 1 To ListView(0).ListItems.Count
            ListView(0).ListItems.Item(X).Selected = True
            If ListView(0).ListItems.Item(X).Checked = True Then
                rsSalvar.Find "codavaliacao=" & "'" & Val(ListView(0).ListItems.Item(X)) & "'"
                If rsSalvar.EOF Then rsSalvar.AddNew
                rsSalvar.Fields(0) = Val(Mid(varGlobal2, 12, 6))
                rsSalvar.Fields(1) = Mid(varGlobal2, 1, 11)
                rsSalvar.Fields(2) = Val(ListView(0).ListItems.Item(X))
                rsSalvar.Fields(3) = ListView(0).SelectedItem.ListSubItems.Item(3)
                rsSalvar.Fields(4) = vCodcoligada ' Codigo da coligada
            End If
        Next
    End If
    If ListView(1).ListItems.Count > 0 Then
        If Not rsSalvar.EOF Then rsSalvar.MoveFirst
        For X = 1 To ListView(1).ListItems.Count
            ListView(1).ListItems.Item(X).Selected = True
            If ListView(1).ListItems.Item(X).Checked = True Then
                rsSalvar.Find "codavaliacao=" & "'" & Val(ListView(1).ListItems.Item(X)) & "'"
                If rsSalvar.EOF Then rsSalvar.AddNew
                rsSalvar.Fields(0) = Val(Mid(varGlobal2, 12, 6))
                rsSalvar.Fields(1) = Mid(varGlobal2, 1, 11)
                rsSalvar.Fields(2) = Val(ListView(1).ListItems.Item(X))
                rsSalvar.Fields(3) = ListView(1).SelectedItem.ListSubItems.Item(3)
                rsSalvar.Fields(4) = vCodcoligada ' Codigo da coligada
            End If
        Next
    End If
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    '>>>>>> GRAVAR OBSERVACAO DA AVALIACAO <<<<<<<<<
    SqlSalvarObs = "select obsavaliacao from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and cpf = '" & Mid(varGlobal2, 1, 11) & "' and codtreinamento = '" & Val(Mid(varGlobal2, 18, 6)) & "' and ativo ='S'"
    rsSalvarObs.Open SqlSalvarObs, cnBanco, adOpenKeyset, adLockOptimistic
    rsSalvarObs.Fields(0) = txtAvaliacao
    If Not rsSalvarObs.EOF Then rsSalvarObs.Update
    rsSalvarObs.Close
    Set rsSalvarObs = Nothing
    cnBanco.CommitTrans
    mobjMsg.Abrir "Os dados da avaliação foram salvos com sucesso", Ok, informacao, "SGC"
    'AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "SGC"
    cnBanco.RollbackTrans
    Exit Sub
TrataErro1:
    Resume Next
End Sub


