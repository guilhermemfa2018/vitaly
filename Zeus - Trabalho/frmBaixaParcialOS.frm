VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBaixaParcialOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa Parcial OS/Operação"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "frmBaixaParcialOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmBaixaParcialOS.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   8760
      Width           =   615
   End
   Begin VB.TextBox txtLvw 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   2040
         OleObjectBlob   =   "frmBaixaParcialOS.frx":1994
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5400
         Top             =   7680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaixaParcialOS.frx":1A02
               Key             =   "Linha"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmBaixaParcialOS.frx":26DC
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   7335
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   12938
         View            =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmBaixaParcialOS.frx":2746
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   8760
      Width           =   615
   End
End
Attribute VB_Name = "frmBaixaParcialOS"
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
'Acima - usado para poder editar o listview --------------------


Private rsBaixaPOS As New ADODB.Recordset
Private sqlBaixaPOS As String

Private Sub cmdCadastro_Click(Index As Integer)
On Error GoTo Err
    Select Case Index
    Case 4
        limpaQualquerDado
        ordenaLVArray ListView2, "5", "0", "3", "6", "7", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbMPBaixaParcial", "idos", "I", Text1
        
        
        Dim rsAtualizaSemana As New ADODB.Recordset
        Dim SqlAtualizaSemana As String
    
        SqlAtualizaSemana = "Update tbMPItens set dataprevista = '" & Format(Date, "dd/mm/yyyy") & "' where idos = '" & Val(Text1.Text) & "'"
        rsAtualizaSemana.Open SqlAtualizaSemana, cnBanco
        
        
        mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "Zeus"
    Case 5
        Unload Me
    End Select
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
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
'On Error GoTo ErrHandler
    Lang_pt_br
    listview_cabecalho
    ResultPesq
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
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Operação", ListView2.Width / 7
    ListView2.ColumnHeaders.Add , , "Nome Op.", ListView2.Width / 3
    
    ListView2.ColumnHeaders.Add , , "Centro de Custo", ListView2.Width / 4
    ListView2.ColumnHeaders.Add , , "% baixado", ListView2.Width / 4.5
    ListView2.ColumnHeaders.Add , , "AlturaLinha", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "idOS", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "revisaoOS", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "datbaixa", ListView2.Width / 10000
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ResultPesq()
    Text1.Text = Mid$(varGlobal, 7, 6)
    LimpaLV ListView2
    Dim vRevisao As Integer
    Dim vChamaSQL As String
    vRevisao = MeuLV.ListView1.SelectedItem.ListSubItems.Item(2)
    Text2.Text = vRevisao
    
    'chamaSQL "select a.idoperacao,a.idcc,b.percentualbaixado,'',a.idos,a.revisaoos from tbMPItens as a left join tbMPBaixaParcial as b on a.idos = b.idos and a.idoperacao = b.idoperacao where a.idos = '" & Val(Text1.Text) & "' and a.revisaoos = '" & vRevisao & "'"
    
    vChamaSQL = vChamaSQL & "SELECT " & vbCrLf
    vChamaSQL = vChamaSQL & " A.IDOPERACAO, A.GRUPO, " & vbCrLf
    vChamaSQL = vChamaSQL & " A.IDCC, " & vbCrLf
    vChamaSQL = vChamaSQL & " B.PERCENTUALBAIXADO, " & vbCrLf
    vChamaSQL = vChamaSQL & " '', " & vbCrLf
    vChamaSQL = vChamaSQL & " A.IDOS, " & vbCrLf
    vChamaSQL = vChamaSQL & " A.REVISAOOS " & vbCrLf
    vChamaSQL = vChamaSQL & "FROM TBMPITENS AS A " & vbCrLf
    vChamaSQL = vChamaSQL & "LEFT JOIN TBMPBAIXAPARCIAL AS B ON " & vbCrLf
    vChamaSQL = vChamaSQL & " A.IDOS = B.IDOS AND " & vbCrLf
    vChamaSQL = vChamaSQL & " A.IDOPERACAO = B.IDOPERACAO " & vbCrLf
    vChamaSQL = vChamaSQL & "WHERE " & vbCrLf
    vChamaSQL = vChamaSQL & " A.IDOS = '" & Val(Text1.Text) & "' AND " & vbCrLf
    vChamaSQL = vChamaSQL & " A.REVISAOOS = '" & vRevisao & "'"
    chamaSQL vChamaSQL
    
    Compoe_Listview2 ListView2, Sqlp, "000"
    ConfLVServ
End Sub

Private Sub ConfLVServ()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems(X).Selected = True
        ListView2.SelectedItem.ListSubItems.Item(3).ReportIcon = "Linha"
    Next
End Sub


'**********************************************
'**********************************************
'**********************************************
'**********************************************
'**********************************************

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView2.SelectedItem Is Nothing Then
        ListView2.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 4 To 4
            leftPos = ListView2.Left + ListView2.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView2.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView2.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView2.SelectedItem.Text
                    Else
                        .Text = ListView2.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    If txtLvw.Enabled = True Then .SetFocus 'Else txtModelo(1).SetFocus
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
        Set lvwCol = ListView2.ColumnHeaders(m_ColIndex)
        
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
            .Left = ListView2.ColumnHeaders.Item(m_ColIndex).Left + 330
            .Top = txtTop + 165
            .Width = 1650
            .Height = 315
        End With
    End If
End Sub

Private Sub txtLvw_GotFocus()
    If txtLvw.Text = "" Then txtLvw.Text = " "
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView2 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
        'other keys can be used for navigation
    End Select
    If txtLvw.Text = "-" Then txtLvw.Text = ""
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
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
        ListView2.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView2.ListItems(ListView2.ListItems.Count) <> c_EntryTxt Then ListView2.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex + 3) = Format(Date, "dd/mm/yyyy")
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > 99 Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = ""
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex) = ""
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
TrataErro:
    Exit Sub
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView2's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView2.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function


