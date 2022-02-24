VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmRelInsp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de relatórios de Inspeção"
   ClientHeight    =   9120
   ClientLeft      =   420
   ClientTop       =   705
   ClientWidth     =   21480
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelInsp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   21480
   Tag             =   "Emissão de relatórios"
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   495
      Left            =   7320
      OleObjectBlob   =   "frmRelInsp.frx":0CCA
      TabIndex        =   35
      Top             =   8520
      Width           =   14055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Siglas"
      Height          =   735
      Left            =   1800
      TabIndex        =   32
      Top             =   8280
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Itens disponíveis para emissão do relatório"
      Height          =   8175
      Left            =   6360
      TabIndex        =   21
      Top             =   120
      Width           =   15015
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "frmRelInsp.frx":0D22
         TabIndex        =   29
         Top             =   7800
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmRelInsp.frx":0D82
         TabIndex        =   28
         Top             =   7800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRelInsp.frx":0DE2
         TabIndex        =   27
         Top             =   7800
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRelInsp.frx":0E50
         TabIndex        =   26
         Top             =   7800
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frmRelInsp.frx":0ECA
         TabIndex        =   25
         Top             =   7800
         Width           =   2535
      End
      Begin VB.TextBox txtLvw 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8880
         TabIndex        =   23
         Top             =   7800
         Width           =   1000
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7455
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList2"
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
      Begin IpToolTips.cIpToolTips cIpToolTips1 
         Left            =   11280
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         BackColor       =   0
      End
      Begin VB.Label Label9 
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   24
         Top             =   7800
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo do Movimento"
      Height          =   735
      Left            =   3840
      TabIndex        =   19
      Top             =   120
      Width           =   2415
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelInsp.frx":0F52
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCadastro 
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
      Index           =   4
      Left            =   120
      Picture         =   "frmRelInsp.frx":0FD2
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "Salvar Relatório"
      ToolTipText     =   "Salvar Relatório"
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
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
      Index           =   6
      Left            =   720
      Picture         =   "frmRelInsp.frx":1C9C
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   8400
      Width           =   615
   End
   Begin VB.Frame Frame8 
      Caption         =   "Data: "
      Height          =   735
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   1695
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   285278209
         CurrentDate     =   40449
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Relatório "
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6135
      Begin VB.Frame Frame4 
         Caption         =   "Inspeções e Teste realizados:"
         Height          =   2415
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   5895
         Begin MSComctlLib.ListView ListView2 
            Height          =   2055
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3625
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   4920
         Width           =   5895
      End
      Begin VB.ComboBox cboCadastro 
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
         ItemData        =   "frmRelInsp.frx":2966
         Left            =   120
         List            =   "frmRelInsp.frx":2973
         TabIndex        =   15
         Tag             =   "Norma de Inspeção"
         ToolTipText     =   "Norma de Inspeção"
         Top             =   1680
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelInsp.frx":299C
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelInsp.frx":2A18
         TabIndex        =   13
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
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
         Height          =   330
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtCadastro 
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
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
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
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
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
         Height          =   330
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelInsp.frx":2A86
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelInsp.frx":2AEC
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmRelInsp.frx":2B5A
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmRelInsp.frx":2BC2
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Relatório nº: "
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   3960
      OleObjectBlob   =   "frmRelInsp.frx":2C2E
      TabIndex        =   34
      Top             =   8880
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmRelInsp"
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
'---------------------------------------------------

'Abaixo ajusta automaticamente a largura das colunas
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
'Acima ajusta automaticamente a largura das colunas
Private X As Integer, W As Integer
Private ContaLV As Integer, LinhaLV As Integer, ContaChecado As Integer, LimiTador As Integer
Private rsLocal As New ADODB.Recordset
Private vPonte1 As TextBox

Private Sub chameleonButton2_Click()
    ChamaGridTrans
    CarregaTipoTrans
    txtCadastro(16).SetFocus
End Sub

'Private Sub chameleonButton3_Click()
'    CompoeListview
'End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 1
'        vPonte1 = Combo1.Text
'        IncluirLV ListView2, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1, vPonte1
    Case 3
'        ExcluirItemLV ListView2
    Case 4
        If cboCadastro(4).Text = "" Then
            mobjMsg.Abrir "Selecione uma Norma de Inspeção", Ok, critico, "Atenção"
            cboCadastro(4).SetFocus
            Exit Sub
        End If
        
        If ContaChecado > LimiTador Then
            mobjMsg.Abrir "Limite máximo de itens selecionados foi ultrapassado." & vbCrLf & "Limite Máximo: " & LimiTador, Ok, critico, "Atenção"
            'Msgbox "Limite máximo de itens selecionados foi ultrapassado." & vbCrLf & "Limite Máximo: " & LimiTador
        Else
            mobjMsg.Abrir "Deseja gravar o relatório?", YesNo, pergunta, "Zeus"
            If Tp = 1 Then
                GravarDados
                Unload Me
            End If
        End If
    Case 6
        mobjMsg.Abrir "Deseja sair da tela de emissão de relatórios?", YesNo, pergunta, "Zeus"
        If Tp = 1 Then
            Unload Me
        End If
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Me.Top = 0
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    
    ContaChecado = 0
    LimiTador = 1000 '49
    'SSTab1.TabEnabled(1) = False
    'optCadastro(0).Value = True
    Legenda = "Aguarde"
    SelecionaLinha
    CompoeControles
'    DesabIcons
    listview_cabecalho 'Chama a Sub que monta o cabeçalho das colunas do Listview
    
    If vSituacao = "INSPEÇÃO DE FABRICAÇÃO" Then
        CompoeListview 'Listview de Fabricação
    Else
        CompoeListview2 'Listview de Pintura
    End If
'    DeterminaPermissão
    
    'CompoeComboSQL Combo1, "Select b.descricao from tbVerifGrupo as a left join tbVerifItem as b on b.codgrupo = a.codgrupo where a.aplicacao = 'Fabricação'"
    
    'Initialize edit box
    txtLvw = ""
    txtLvw.Visible = False
    txtLvw.Tag = False 'is ListView2 dirty, not used in this example
    'HookEdtLvw ListView1.hWnd 'subclass scroll events
    SkinLabel11.Caption = SkinLabel11.Caption & SkinLabel7.Caption
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delas e e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Posição", ListView1.Width / 13
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Desenho", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Rev.", ListView1.Width / 22
    ListView1.ColumnHeaders.Add , , "Q. Total", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Peso", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Q. Pendente", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Q. à lib.", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "codFase", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "UN", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "CodLM", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "CodSeq", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Peso Lib.", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Insp. Realizadas", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Possui Pintura?", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "OS", ListView1.Width / 10000
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID. Insp.", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Descrição", ListView2.Width / 1.5
    ListView2.ColumnHeaders.Add , , "Sigla", ListView2.Width / 9
    
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(13).Alignment = lvwColumnRight
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub CompoeCombo3()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Dim X As Integer, Y As Integer
    Dim fase As String
    Y = ListView1.ListItems.Count
    X = 1
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 1
    Me.ListView1.SortOrder = lvwAscending
    
    ListView1.ListItems(X).Selected = True
    fase = ListView1.SelectedItem.ListSubItems.Item(1)
    cboCadastro(3).AddItem fase
    For X = 2 To Y
        ListView1.ListItems(X).Selected = True
        If fase <> ListView1.SelectedItem.ListSubItems.Item(1) Then
            fase = ListView1.SelectedItem.ListSubItems.Item(1)
            cboCadastro(3).AddItem fase
        End If
    Next
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim rsRelInsp As New ADODB.Recordset
    Dim sqlRelInsp As String
    
    sqlRelInsp = "select a.codprojeto,a.fce,a.projeto,c.nome from tbProjetos as a inner join tbFO as b on a.fce = b.fce inner join tbclifor as c on b.codclifor = c.codclifor " & _
                 "where a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and a.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' Order by a.fce desc,a.descricao"
    rsRelInsp.Open sqlRelInsp, cnBanco, adOpenKeyset, adLockReadOnly
    If rsRelInsp.RecordCount = 0 Then Exit Sub
    
    txtCadastro(4) = Format(GeraCodigo, "000000000") & "" 'Identificador do relatório
    txtCadastro(0) = Val(Mid(varGlobal, 7, 4)) 'FCE nº
    txtCadastro(1) = rsRelInsp.Fields(0) 'ID Projeto
    txtCadastro(2) = rsRelInsp.Fields(2) 'Descrição do projeto
    txtCadastro(3) = rsRelInsp.Fields(3) 'Nome do cliente
    DTPicker1 = Date 'Data de emissão do relatório
    varGlobal2 = txtCadastro(4).Text
    SkinLabel7.Caption = vSituacao
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

Private Sub CompoeListview()
On Error GoTo Err
    Dim rsLisview As New ADODB.Recordset
    Dim sqlLisview As String
    Dim W As Integer
    
    sqlLisview = "select T1.fce,max(T1.codprojeto) as codprojeto,T1.projeto,max(T1.codlm) codlm,max(T1.codseq) codseq,T1.desenho,max(T1.revisao) as revisao,T1.descricao,T1.posicao,max(T1.quantidade) as quantidade,max(T1.PesoPosicao) as PesoPosicao,case when sum(T1.qtdlib) is null then 0 else sum(T1.qtdlib) end as qtdlib,case when MAX(T1.quantcj)- sum(T1.qtdlib) is null then MAX(T1.quantcj) else MAX(T1.quantcj)- sum(T1.qtdlib) end as qtddisp,max(T1.pintura) as pintura " & _
                 "from (Select a.fce,d.codprojeto,d.projeto,a.codlm,a.codseq,c.desenho,c.revisao,b.descposicao as descricao,b.posicao as posicao,a.quantcj as quantidade,b.pesoposicao,f.qtdlib,a.quantcj,a.area as pintura From tbItemLM as a inner join tbPosicoes as b on a.codigopos = b.codigopos inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbProjetos as d on c.codprojeto = d.codprojeto left join tbrelinspexpitens as f on a.fce = f.fce and " & _
                 "d.codprojeto = f.codprojeto and a.codlm = f.codlm and a.codseq = f.codseq) T1 where T1.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and T1.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' group by T1.fce,T1.projeto,T1.desenho,T1.posicao,T1.descricao order by T1.fce,T1.projeto,T1.desenho,T1.posicao"
    
    
    cnBanco.CommandTimeout = 0
    rsLisview.Open sqlLisview, cnBanco, adOpenKeyset, adLockReadOnly
    If rsLisview.RecordCount <> 0 Then Principal.ProgressBar1.Max = rsLisview.RecordCount
    W = 0
    ListView1.ListItems.Clear
    Do While Not rsLisview.EOF
        Principal.ProgressBar1.Value = W
        If rsLisview.Fields(12) > 0 Then
            Set ItemLst = ListView1.ListItems.Add(, , rsLisview.Fields(8)) 'Posição identificador
            ItemLst.SubItems(1) = rsLisview.Fields(7) 'Posição descrição
            ItemLst.SubItems(2) = rsLisview.Fields(5) 'Desenho
            ItemLst.SubItems(3) = rsLisview.Fields(6) 'Revisão
            ItemLst.SubItems(4) = rsLisview.Fields(9) 'Quantidade total
            ItemLst.SubItems(5) = Format(rsLisview.Fields(10), "#,##0.00;(#,##0.00)") 'Peso Posição
            ItemLst.SubItems(6) = rsLisview.Fields(12) 'Q. Pendente
            ItemLst.SubItems(7) = " " 'Quantidade à liberar
            If rsLisview.Fields(13) > 0 Then
                ItemLst.SubItems(8) = "3" 'Código da Fase de (Liberação de fabricação)
            Else
                ItemLst.SubItems(8) = "10" 'Pula para fase de expedição se não houver pintura
            End If
            
            ItemLst.SubItems(9) = "CJ" 'Unidade de medida
            ItemLst.SubItems(10) = rsLisview.Fields(3) 'Código da LM - Lista de Materiais
            ItemLst.SubItems(11) = rsLisview.Fields(4) 'Código da sequência da LM
            ItemLst.SubItems(12) = " " 'Peso Liberado da Posição
            ItemLst.SubItems(13) = " " 'Inspeções realizadas
            If rsLisview.Fields(13) > 0 Then
                ItemLst.SubItems(14) = "S" 'Possui Pintura? (S/N)
            Else
                ItemLst.SubItems(14) = "N" 'Possui Pintura? (S/N)
            End If
            
            Dim rsAchaOS As New ADODB.Recordset
            Dim sqlAchaOS As String
            sqlAchaOS = "SELECT A.idoperacao,A.idcc,A.status,A.idos FROM TBOSITENS AS A WHERE A.FCE = '" & Val(Mid(varGlobal, 7, 4)) & "' AND A.codlm = '" & rsLisview.Fields(3) & "' AND A.codseq = '" & rsLisview.Fields(4) & "' AND A.idcc IN('7000.7103.SC-02','7000.7103.SC-03') "
            cnBanco.CommandTimeout = 0
            rsAchaOS.Open sqlAchaOS, cnBanco, adOpenKeyset, adLockReadOnly
            If rsAchaOS.RecordCount <> 0 Then
                ItemLst.SubItems(15) = rsAchaOS.Fields(3) 'rsLisview.Fields(17) 'OS Nº
            Else
                ItemLst.SubItems(15) = "-" 'rsLisview.Fields(17) 'OS Nº
            End If
            rsAchaOS.Close
            Set rsAchaOS = Nothing
        End If
        
        W = W + 1
        If Not rsLisview.EOF Then rsLisview.MoveNext
    Loop
    Principal.ProgressBar1.Value = 0
    rsLisview.Close
    Set rsLisview = Nothing
    
'    Select b.descricao from tbVerifGrupo as a left join tbVerifItem as b on b.codgrupo = a.codgrupo where a.aplicacao = 'Fabricação'
'AINDA NÃO ESTA FILTRANDO POR FCE.
'AO TERMINAR AS ROTINAS DOS RELATÓRIO DEVERA SER CONFIGURADO PARA FILTRAR DE ACORDO COM O QUE FOI INFORMADO PARA A FCE NO SETOR DE PLANEJAMENTO
    ListView2.ListItems.Clear
'    chamaSQL "Select right('00000' + rtrim(a.codgrupo),2) + right('00000' + rtrim(b.coditem),2),b.descricao,b.sigla from tbVerifGrupo as a left join tbVerifItem as b on b.codgrupo = a.codgrupo where a.aplicacao = 'Fabricação'"
    chamaSQL "select right('00000' + rtrim(a.codgrupo),2) + right('00000' + rtrim(a.coditem),2), c.descricao,c.sigla from tbListaVerif as a inner join tbVerifGrupo as b on b.codgrupo = a.codgrupo inner join tbVerifItem as c on c.codgrupo = a.codgrupo and c.coditem = a.coditem where b.aplicacao = 'Fabricação' and a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "'"
    Compoe_Listview ListView2, Sqlp, "00"
    MarcaDesmarca ListView2
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

Private Sub CompoeListview2()
On Error GoTo Err
    Dim rsLisview As New ADODB.Recordset
    Dim sqlLisview As String
    
    Dim rsLisview2 As New ADODB.Recordset
    Dim sqlLisview2 As String
    Dim W As Integer
    sqlLisview = "Select a.fce,a.codprojeto,a.codlm,a.codseq,max(a.desenho) as desenho,max(a.revisao) as revisao,max(a.descposicao) as descposicao,max(a.posicao) as posicao,sum(a.pesolib) as pesolib,max(a.status) as status,sum(a.qtdlib) as qtddisp  " & _
                 "from tbrelinspexpitens as a where a.status = 3 and a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and a.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' group by a.fce,a.codprojeto,a.codlm,a.codseq"
    rsLisview.Open sqlLisview, cnBanco, adOpenKeyset, adLockReadOnly
    If rsLisview.RecordCount <> 0 Then Principal.ProgressBar1.Max = rsLisview.RecordCount
    W = 0
    ListView1.ListItems.Clear
    Do While Not rsLisview.EOF
        Principal.ProgressBar1.Value = W
        sqlLisview2 = "Select a.fce,a.codprojeto,a.codlm,a.codseq,max(a.desenho) as desenho,max(a.revisao) as revisao,max(a.descposicao) as descposicao,max(a.posicao) as posicao,sum(a.pesolib) as pesolib,max(a.status) as status,sum(a.qtdlib) as qtddisp  " & _
                     "from tbrelinspexpitens as a where a.status = 10 and a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and a.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' and a.codlm = '" & rsLisview.Fields(2) & "' and a.codseq ='" & rsLisview.Fields(3) & "' group by a.fce,a.codprojeto,a.codlm,a.codseq "
        rsLisview2.Open sqlLisview2, cnBanco, adOpenKeyset, adLockReadOnly
        
        If rsLisview2.RecordCount > 0 Then
            'ENTRA NESSA CONDIÇÃO SE O RESULTADO DA CONSULTA rsLisview2 NÃO FOR VAZIA
            'E SOMENTE ENTRA NA CONDIÇÃO ABAIXO SE A POSIÇÃO 10 DA PRIMEIRA CONSULTA - A POSIÇÃO 10 DA SEGUNDA CONSULTA FOR MAIOR QUE ZERO
            If rsLisview.Fields(10) - rsLisview2.Fields(10) > 0 Then
                Set ItemLst = ListView1.ListItems.Add(, , rsLisview.Fields(7)) 'Posição identificador
                ItemLst.SubItems(1) = rsLisview.Fields(6) 'Posição descrição
                ItemLst.SubItems(2) = rsLisview.Fields(4) 'Desenho
                ItemLst.SubItems(3) = rsLisview.Fields(5) 'Revisão
                ItemLst.SubItems(4) = rsLisview.Fields(10) 'Quantidade total
                ItemLst.SubItems(5) = Format(rsLisview.Fields(8), "#,##0.00;(#,##0.00)") 'Peso Posição
                
                ItemLst.SubItems(6) = rsLisview.Fields(10) - rsLisview2.Fields(10) 'Q. Pendente
                
                ItemLst.SubItems(7) = " " 'Quantidade à liberar
                ItemLst.SubItems(8) = "10" 'Código da Fase de (Liberação de Pintura)
                ItemLst.SubItems(9) = "CJ" 'Unidade de medida
                ItemLst.SubItems(10) = rsLisview.Fields(2) 'Código da LM - Lista de Materiais
                ItemLst.SubItems(11) = rsLisview.Fields(3) 'Código da sequência da LM
                ItemLst.SubItems(12) = " " 'Peso Liberado da Posição
                ItemLst.SubItems(13) = " " 'Inspeções realizadas
            End If
        Else
            'ENTRA NESSA CONDIÇÃO SE O RESULTADO DA CONSULTA rsLisview2 FOR VAZIA
            Set ItemLst = ListView1.ListItems.Add(, , rsLisview.Fields(7)) 'Posição identificador
            ItemLst.SubItems(1) = rsLisview.Fields(6) 'Posição descrição
            ItemLst.SubItems(2) = rsLisview.Fields(4) 'Desenho
            ItemLst.SubItems(3) = rsLisview.Fields(5) 'Revisão
            ItemLst.SubItems(4) = rsLisview.Fields(10) 'Quantidade total
            ItemLst.SubItems(5) = Format(rsLisview.Fields(8), "#,##0.00;(#,##0.00)") 'Peso Posição
            ItemLst.SubItems(6) = rsLisview.Fields(10) 'Q. Pendente
            ItemLst.SubItems(7) = " " 'Quantidade à liberar
            ItemLst.SubItems(8) = "10" 'Código da Fase de (Liberação de fabricação)
            ItemLst.SubItems(9) = "-" 'Unidade de medida
            ItemLst.SubItems(10) = rsLisview.Fields(2) 'Código da LM - Lista de Materiais
            ItemLst.SubItems(11) = rsLisview.Fields(3) 'Código da sequência da LM
            ItemLst.SubItems(12) = " " 'Peso Liberado da Posição
            ItemLst.SubItems(13) = " " 'Inspeções realizadas
        End If
        rsLisview2.Close
        W = W + 1
        If Not rsLisview.EOF Then rsLisview.MoveNext
    Loop
    Principal.ProgressBar1.Value = 0
    rsLisview.Close
    Set rsLisview = Nothing
    Set rsLisview2 = Nothing
    
    ListView2.ListItems.Clear
'    chamaSQL "Select right('00000' + rtrim(a.codgrupo),2) + right('00000' + rtrim(b.coditem),2),b.descricao,b.sigla from tbVerifGrupo as a left join tbVerifItem as b on b.codgrupo = a.codgrupo where a.aplicacao = 'Pintura'"
    chamaSQL "select right('00000' + rtrim(a.codgrupo),2) + right('00000' + rtrim(a.coditem),2), c.descricao,c.sigla from tbListaVerif as a inner join tbVerifGrupo as b on b.codgrupo = a.codgrupo inner join tbVerifItem as c on c.codgrupo = a.codgrupo and c.coditem = a.coditem where b.aplicacao = 'Pintura' and a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "'"
    Compoe_Listview ListView2, Sqlp, "00"
    MarcaDesmarca ListView2
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

Private Sub compoeSiglas()
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    txtCadastro(6).Text = ""
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems(X).Selected = True
        If ListView2.ListItems.Item(X).Checked = True Then
            If txtCadastro(6).Text = "" Then
                txtCadastro(6).Text = ListView2.SelectedItem.ListSubItems.Item(2)
            Else
                txtCadastro(6).Text = txtCadastro(6).Text & "/" & ListView2.SelectedItem.ListSubItems.Item(2)
            End If
        End If
    Next
End Sub

Private Sub AlteraListview()
    On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    Contador = 0
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    varGlobal = ListView1.ListItems.Item(X)
    Exit Sub
Err:
    Msgbox "Nenhuma Ficha de Orçamento selecionada", vbInformation, "ZEUS"
    Exit Sub
End Sub

Private Sub GravarDados()
On Error GoTo Err

    Dim Y As Integer, X As Integer
    
    Dim rsRelatorio As New ADODB.Recordset
    Dim sqlRelatorio As String
    Dim rsItensRelatorio As New ADODB.Recordset
    Dim sqlItensRelatorio As String
    
    
    If SomaTotais = False Then Exit Sub
    If Val(SkinLabel29) = 0 Then
        mobjMsg.Abrir "Nenhum Item Selecionado para compor o relatório", Ok, critico, "Atenção"
'       Msgbox "Nenhum Item Selecionado para compor o relatório", vbCritical, "Zeus"
        Exit Sub
    End If
    
10  cnBanco.BeginTrans

    txtCadastro(4) = Format(GeraCodigo, "000000000") & "" 'Identificador do relatório

    sqlRelatorio = "select * from tbRelInspExp"
    rsRelatorio.Open sqlRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    rsRelatorio.AddNew
    rsRelatorio.Fields(0) = Val(txtCadastro(4)) 'Codigo do Relatorio
    rsRelatorio.Fields(1) = Val(txtCadastro(0)) 'FCE
    rsRelatorio.Fields(2) = Val(txtCadastro(1)) 'Codigo do projeto
    rsRelatorio.Fields(3) = Format(DTPicker1, "dd/mm/yyyy") 'Data do relatorio
    rsRelatorio.Fields(4) = txtCadastro(5) 'Observação
    rsRelatorio.Fields(5) = 0 'Status de impressão
    rsRelatorio.Fields(6) = cboCadastro(4) 'Norma de Liberação
    If SkinLabel7.Caption = "INSPEÇÃO DE FABRICAÇÃO" Then
        rsRelatorio.Fields(7) = 3 'Tipo relatorio (3) Inspeção de Fabrica
    Else
        rsRelatorio.Fields(7) = 10 'Tipo relatorio (10) Inspeção de Pintura
    End If
    rsRelatorio.Fields(9) = NomUsu
    rsRelatorio.Update
    rsRelatorio.Close
    Set rsRelatorio = Nothing


    'Gravar dados referente aos Itens do Relatório
    sqlItensRelatorio = "select * from tbRelInspExpitens"
    rsItensRelatorio.Open sqlItensRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If ListView1.ListItems.Item(X).Checked = True Then
            rsItensRelatorio.AddNew
            rsItensRelatorio.Fields(0) = Val(txtCadastro(4)) 'Codigo do relatorio
            rsItensRelatorio.Fields(1) = Val(txtCadastro(0).Text) 'Nº FCE
            rsItensRelatorio.Fields(2) = Val(txtCadastro(1)) 'Código do Projeto
            rsItensRelatorio.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2) 'Desenho
            rsItensRelatorio.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(3) 'Revisão do Desenho
            rsItensRelatorio.Fields(5) = ListView1.ListItems.Item(X) 'Posição
            rsItensRelatorio.Fields(6) = ListView1.SelectedItem.ListSubItems.Item(1) 'Descrição da posição
            If SkinLabel7.Caption = "INSPEÇÃO DE FABRICAÇÃO" Then
                rsItensRelatorio.Fields(7) = 3 'ListView1.SelectedItem.ListSubItems.Item(8) 'Status (Codfase)
            Else
                rsItensRelatorio.Fields(7) = 10 'ListView1.SelectedItem.ListSubItems.Item(8) 'Status (Codfase)
            End If
            rsItensRelatorio.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(7) 'Quantidade liberada
            rsItensRelatorio.Fields(9) = ListView1.SelectedItem.ListSubItems.Item(12) 'Peso liberada
            rsItensRelatorio.Fields(10) = Val(ListView1.SelectedItem.ListSubItems.Item(10)) 'Código da LM - Lista de Material
            rsItensRelatorio.Fields(11) = Val(ListView1.SelectedItem.ListSubItems.Item(11)) 'Código da sequencia da LM
            rsItensRelatorio.Fields(13) = ListView1.SelectedItem.ListSubItems.Item(13) 'Inspeções realizadas
            rsItensRelatorio.Update
        End If
    Next
    rsItensRelatorio.Close
    Set rsItemRelatorio = Nothing
    
   
    cnBanco.CommitTrans
    mobjMsg.Abrir "Dados gravados com sucesso", Ok, informacao, "Atenção"
    
    mobjMsg.Abrir "Dados gravados com sucesso.Deseja imprimir de relatório?", YesNo, pergunta, "Zeus"
    If Tp = 1 Then
        vCodRel = Val(txtCadastro(4))
        Dim rsInspecao As New ADODB.Recordset
        Dim sqlInspecao As String
        limpaQualquerDado
        
        sqlInspecao = "select b.descricao,b.sigla from tbVerifGrupo as a inner join tbVerifItem as b on a.codgrupo = b.codgrupo where a.aplicacao <> '-'"
        rsInspecao.Open sqlInspecao, cnBanco, adOpenKeyset, adLockReadOnly
        Y = rsInspecao.RecordCount
        For X = 1 To Y
            vQualquerDado(0, X) = rsInspecao.Fields(1) & " - " & rsInspecao.Fields(0)
            rsInspecao.MoveNext
        Next
        rsInspecao.Close
        Set rsInspecao = Nothing
        vQualquerDado(0, 30) = SkinLabel11.Caption
        FCRLibFab.Show 1
        sqlRelatorio = "update tbRelInspExp set statusimp=1 where codrel ='" & vCodRel & "'"
        rsRelatorio.Open sqlRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    End If
    'CompoeControles
    'CompoeListview
    Unload Me
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
    Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub SelecionaLinha()
On Error GoTo Err
    Dim Y As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For W = 1 To Y
        If MeuLV.ListView1.ListItems.Item(W).Selected = True Then
            Exit For
        End If
    Next
    MeuLV.ListView1.ListItems(W).Selected = True
Err:
    Resume Next
End Sub

Private Function VerFaseDiferente()
'    VerFaseDiferente = True
'    Dim X As Integer, Y As Integer
'    Dim fase As String
'    Y = ListView1.ListItems.Count
'    fase = ""
'    For X = 1 To Y
'        If ListView1.ListItems.Item(X).Checked = True Then
'            ListView1.ListItems(X).Selected = True
'            If fase = "" Then fase = ListView1.SelectedItem.ListSubItems.Item(1)
'
'            If optCadastro(1).Value = True And UCase(ListView1.SelectedItem.ListSubItems.Item(1)) <> UCase("Expedição") Or optCadastro(1).Value = False And UCase(ListView1.SelectedItem.ListSubItems.Item(1)) = UCase("Expedição") Then
'                Msgbox "Selecione apenas itens referentes ao tipo de relatório escolhido", vbCritical, "Zeus"
'                VerFaseDiferente = False
'                Exit For
'            End If
'
'            If fase <> ListView1.SelectedItem.ListSubItems.Item(1) Then
'                Msgbox "Não é permitido selecionar fases diferentes", vbCritical, "Zeus"
'                VerFaseDiferente = False
'                Exit For
'            End If
'            If Label9 <> "ADMINISTRADOR" Then
'                If Label9 <> "EXPEDIÇÃO" And ListView1.SelectedItem.ListSubItems.Item(1) = "Expedição" Then
'                    Msgbox "O Usuário não tem permissão para liberar essa fase", vbCritical, "Zeus"
'                    VerFaseDiferente = False
'                    Exit For
'                End If
'                If Label9 = "EXPEDIÇÃO" And ListView1.SelectedItem.ListSubItems.Item(1) <> "Expedição" Then
'                    Msgbox "O Usuário não tem permissão para liberar essa fase", vbCritical, "Zeus"
'                    VerFaseDiferente = False
'                    Exit For
'                End If
'            End If
'        End If
'    Next
'    Me.ListView1.SortKey = 0
'    Me.ListView1.SortOrder = lvwAscending
End Function

Private Sub DeterminaPermissão()
'    If VEXP <> 0 And VINS <> 0 Then
'        Label9 = "ADMINISTRADOR"
'    ElseIf VEXP <> 0 And VINS = 0 Then
'        Label9 = "EXPEDIÇÃO"
'        optCadastro(1).Value = True
'        optCadastro(0).Enabled = False
'        optCadastro(1).Enabled = False
'    ElseIf VEXP = 0 And VINS <> 0 Then
'        Label9 = "INSPEÇÃO"
'        optCadastro(0).Value = True
'        optCadastro(0).Enabled = False
'        optCadastro(1).Enabled = False
'    ElseIf VEXP = 0 And VINS = 0 Then
'        Label9 = "VISUALIZAR"
'        optCadastro(0).Enabled = False
'        optCadastro(1).Enabled = False
'    End If
End Sub

Private Sub PosLinha()
    Dim ContaLV As Integer
    ContaLV = ListView1.ListItems.Count
    For LinhaLV = 1 To ContaLV
        If ListView1.ListItems.Item(LinhaLV).Selected = True Then
            Exit For
        End If
    Next
End Sub

Private Function SomaTotais()
On Error GoTo TrataErro
    SkinLabel12.Caption = ""
    SomaTotais = True
    Dim Y As Integer, SomaPeso As Double, SomaQtd As Double
    Y = ListView1.ListItems.Count
    SomaQtd = 0
    SomaPeso = 0
    For W = 1 To Y
        If ListView1.ListItems.Item(W).Checked = True Then
            ListView1.ListItems(W).Selected = True
            SomaQtd = SomaQtd + ListView1.SelectedItem.ListSubItems.Item(7)
            SomaPeso = SomaPeso + ((ListView1.SelectedItem.ListSubItems.Item(5) / ListView1.SelectedItem.ListSubItems.Item(4)) * ListView1.SelectedItem.ListSubItems.Item(7))
            ListView1.SelectedItem.ListSubItems.Item(12) = Format(((ListView1.SelectedItem.ListSubItems.Item(5) / ListView1.SelectedItem.ListSubItems.Item(4)) * ListView1.SelectedItem.ListSubItems.Item(7)), "#,##0.00;(#,##0.00)")
        End If
    Next
    SkinLabel29 = Format(SomaQtd, "#,##0.00;(#,##0.00)")
    SkinLabel30 = Format(SomaPeso, "#,##0.00;(#,##0.00)")
    Exit Function
TrataErro:
    SomaTotais = False
    'mobjMsg.Abrir "Existem itens marcados no Relatorio que não possuem Qtd. liberada", Ok, informacao, "Atenção"
    SkinLabel12.Caption = "Existem itens marcados no Relatorio que não possuem Qtd. liberada"
End Function

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tbRelInspExp order by codrel Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraCodigo.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        'QualForm = "novalm"
        If IniciaRelsEm > 0 Then
            GeraCodigo = IniciaRelsEm
        Else
            GeraCodigo = 1 'NovoCodigo
        End If
    End If
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
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

Private Sub ChamaGridTrans()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTransportadoras order by nome"
    procnom = "nome"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Transportadoras"
    Pesquisa = frmRelInsp.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nome=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadastro(6).Text = Format(rsLocal.Fields(0), "000000")
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

Private Sub CarregaTipoTrans()
On Error GoTo Err
    Dim X As Integer
    Dim rsTipoTrans As New ADODB.Recordset
    SqlM = "Select * from tbTransportadoras order by tbTransportadoras.codtransp"
    rsTipoTrans.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsTipoTrans.EOF Then rsTipoTrans.MoveFirst
    rsTipoTrans.Find "codtransp=" & "'" & Val(Me.txtCadastro(6)) & "'"
    If rsTipoTrans.EOF Then
        txtCadastro(6).Text = Format(txtCadastro(6), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            Msgbox "Transportadora não cadastrada", vbInformation, "Zeus"
            txtCadastro(7) = ""
        End If
    Else
        txtCadastro(6).Text = Format(rsTipoTrans.Fields(0), "000000") & ""
        txtCadastro(7).Text = rsTipoTrans.Fields(1)
        txtCadastro(8).Text = rsTipoTrans.Fields(2)
        txtCadastro(9).Text = rsTipoTrans.Fields(3)
        txtCadastro(10).Text = rsTipoTrans.Fields(4)
        txtCadastro(11).Text = rsTipoTrans.Fields(5)
        txtCadastro(12).Text = rsTipoTrans.Fields(6)
        txtCadastro(13).Text = rsTipoTrans.Fields(7)
        cboCadastro(0).Text = rsTipoTrans.Fields(8)
        For X = 7 To 13
            txtCadastro(X).Enabled = False
        Next
        cboCadastro(0).Enabled = False
    End If
    rsTipoTrans.Close
    Set rsTipoTrans = Nothing
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

Private Sub Form_Resize()
    DimensionaFormInsp
End Sub

Private Sub ListView1_Click()
    Dim M As Integer
    M = ListView1.ListItems.Count
    PosLinha
    ContaChecado = 0
    SkinLabel10 = "Registros selecionadas: "
    compoeSiglas
    For K = 1 To M
        If ListView1.ListItems.Item(K).Selected = True Then
            If ListView1.ListItems.Item(K).Checked = False Then
                ListView1.ListItems.Item(K).Checked = True
                ContaChecado = ContaChecado + 1
                ListView1.SelectedItem.ListSubItems.Item(13) = txtCadastro(6).Text
                If ListView1.SelectedItem.ListSubItems.Item(7) = " " And txtLvw = "" Then
                    ListView1.SelectedItem.ListSubItems.Item(7) = ListView1.SelectedItem.ListSubItems.Item(6)
                End If
                If Val(ListView1.SelectedItem.ListSubItems.Item(7)) > Val(ListView1.SelectedItem.ListSubItems.Item(6)) Then ListView1.SelectedItem.ListSubItems.Item(7) = ListView1.SelectedItem.ListSubItems.Item(6)
            Else
                ListView1.ListItems.Item(K).Checked = False
                'ListView1.SelectedItem.ListSubItems.Item(7) = " "
                If txtLvw = "" Or ListView1.SelectedItem.ListSubItems.Item(7) = " " Then
                    ListView1.SelectedItem.ListSubItems.Item(12) = " "
                    ListView1.SelectedItem.ListSubItems.Item(13) = " "
                    txtLvw = ""
                End If
            End If
        Else
            If ListView1.ListItems.Item(K).Checked = True Then
                ContaChecado = ContaChecado + 1
            End If
        End If
    Next
    SomaTotais
    SkinLabel10 = SkinLabel10 & ContaChecado
    If ListView1.ListItems.Count > 0 Then
        ListView1.ListItems(LinhaLV).Selected = True
    End If
End Sub

'Private Sub optCadastro_Click(Index As Integer)
'    Select Case Index
'    Case 0 'INSPEÇÃO
'        SSTab1.TabEnabled(1) = False
'        cboCadastro(4).Enabled = True
'        LimiTador = 1000 '49
'    Case 1 'EXPEDIÇÃO
'        SSTab1.TabEnabled(1) = True
'        cboCadastro(4).Enabled = False
'        LimiTador = 1000 '37
'    End Select
'End Sub
'---------------------------
'---------------------------
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'fired when the listitem is already selected, for this reason can't used mousedown event
'so we know which row is clicked, for the column, we need to translate the x to listview coordinate
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 8 To 8
            PosLinha
            If ListView1.ListItems.Item(LinhaLV).Checked = False Then Exit Sub
            
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
                    .SetFocus
                End With
                Exit For
            End If
        Next i
    End If
End If
End Sub

Function GetLvwDeltaX() As Single
'returns deltaX, the scroll distance in pixels relative to ListView2.left, how much we scroll right
'si.npage propotional to both the width of the scroll box and ListView2.width
'si.npos is the scrolling position, which is propotional to deltaX

    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView1.ColumnHeaders(ListView1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    'PrintLvwColInfo
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView1.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    '58 is some constant to get things just right
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
'called from ListView2 mouseup and subclass scroll events
'constants used are determined by trial & error, these are mainly the various widths and heights
'of edges in the classical windows. these constants may not be correct for other windows styles.
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
        .Left = ListView1.SelectedItem.Left + 10295 '11795
        .Top = txtTop
        '.Width =
        .Height = ListView1.SelectedItem.Height - 8
    End With
End If
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    If Index = 6 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaTipoTrans
        End If
    End If
Error:
    Exit Sub
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView2 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
            If ListView1.SelectedItem.ListSubItems.Item(7) > ListView1.SelectedItem.ListSubItems.Item(6) Then ListView1.SelectedItem.ListSubItems.Item(7) = ListView1.SelectedItem.ListSubItems.Item(6)
            If ListView1.ListItems.Item(LinhaLV).Checked = False Then ListView1.SelectedItem.ListSubItems.Item(7) = ""
        'other keys can be used for navigation
    End Select
End Sub

Private Sub txtLvw_LostFocus()
    If m_ColIndex = 1 Then
        ListView1.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
    ElseIf m_ColIndex Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView2's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

'FUNCAO PARA MUDAR TOOLTIPS
Private Sub MudaTool()
    On Error Resume Next
    Dim Ctl As Control
    Dim i As Integer
    With Me.cIpToolTips1
        .Create
        .Title = "Atenção:" 'Titulo do tooltip
        .MyIcon = itInfoIcon 'Icone do tooltip
        .BackColor = &H80000018  'Cor de fundo
        .ForeColor = &H800000    'Cor da letra e bordas
        For Each Ctl In Me.Controls
            If Ctl.Tag <> "" Then
                .AddTool Ctl, tfAbsolute, Replace(Ctl.Tag, "|", vbCrLf)
            End If
        Next
    End With
End Sub

Private Sub CompoeTabTemp()
On Error GoTo Err
    Dim rsCodRel As New ADODB.Recordset
    Dim rsTbTemp As New ADODB.Recordset
    Dim sqlCodRel As String, sqlTbTemp As String
    Dim Y As Integer, X As Integer
    
    sqlTbTemp = "Delete from tbtemp"
    rsTbTemp.Open sqlTbTemp, cnBanco
    
    sqlTbTemp = "Select * from tbtemp"
    rsTbTemp.Open sqlTbTemp, cnBanco, adOpenKeyset, adLockOptimistic
    
    Y = frmRelInsp.ListView1.ListItems.Count
    For X = 1 To Y
        frmRelInsp.ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If frmRelInsp.ListView1.ListItems.Item(X).Checked = True Then
            sqlCodRel = "select tbrelatorios.codrel,tbitemrelatorio.idld from tbitemrelatorio inner join tbRelatorios on tbrelatorios.codrel = tbitemrelatorio.codrel where tbitemrelatorio.idld = '" & Val(frmRelInsp.ListView1.ListItems.Item(X)) & "'" & " order by tbitemrelatorio.codprocesso,tbitemrelatorio.codfase desc"
            rsCodRel.Open sqlCodRel, cnBanco, adOpenKeyset, adLockReadOnly
            rsCodRel.MoveNext
            If rsCodRel.RecordCount > 0 Then
                rsTbTemp.AddNew
                rsTbTemp.Fields(0) = rsCodRel.Fields(0)
                rsTbTemp.Fields(1) = rsCodRel.Fields(1)
            End If
            rsCodRel.Close
        End If
    Next
    Set rsCodRel = Nothing
    rsTbTemp.Update
    Set rsTbTemp = Nothing
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
