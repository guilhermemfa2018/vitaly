VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmAtividades 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Movimentações - OS"
   ClientHeight    =   4065
   ClientLeft      =   2160
   ClientTop       =   4020
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAtividades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
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
      Height          =   615
      Left            =   6600
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "Status da Movimentação"
         ToolTipText     =   "Status da Movimentação"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   1095
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "Descrição da Movimentação"
         ToolTipText     =   "Descrição da Movimentação"
         Top             =   1920
         Width           =   7335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAtividades.frx":0CCA
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox cboCadEscolaridade 
         Height          =   345
         ItemData        =   "frmAtividades.frx":0D36
         Left            =   1320
         List            =   "frmAtividades.frx":0D43
         TabIndex        =   1
         Text            =   "Parada"
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Tag             =   "Código da Movimentação"
         ToolTipText     =   "Código da Movimentação"
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAtividades.frx":0D62
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtCadEscolaridade 
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Identificador automático"
         ToolTipText     =   "Identificador automático"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Tag             =   "Nome da Movimentação"
         ToolTipText     =   "Nome da Movimentação"
         Top             =   1200
         Width           =   6135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmAtividades.frx":0DC8
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAtividades.frx":0E2A
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   3360
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
      MICON           =   "frmAtividades.frx":0E88
      PICN            =   "frmAtividades.frx":0EA4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "Salvar registro"
      ToolTipText     =   "Salvar registro"
      Top             =   3360
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
      MICON           =   "frmAtividades.frx":1B7E
      PICN            =   "frmAtividades.frx":1B9A
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
Attribute VB_Name = "frmAtividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsParadas As New ADODB.Recordset
Private sqlParadas As String
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados da Movimentação?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravarDados
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Movimentações - OS?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Unload Me
            Set frmAtividades = Nothing
        End If
    End Select
End Sub

Private Sub cmdCadastro_MouseOver(Index As Integer)
    Legenda = cmdCadastro(Index).ToolTipText
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        DesbloqueiaControles
    End If
    configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub GravarDados()
On Error GoTo Err
    If ValidaCampo = False Then Exit Sub
    Dim rsParadas As New ADODB.Recordset
    Dim sqlParadas As String
    Dim Y As Integer
'    cnBanco.BeginTrans
   
    sqlParadas = "select * from tbparadas where idparada = '" & txtCadEscolaridade(0) & "'"
    rsParadas.Open sqlParadas, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsParadas.EOF Then rsParadas.AddNew
    'rsParadas.Fields(0) = Val(txtCadEscolaridade(0))
    rsParadas.Fields(1) = cboCadEscolaridade
    rsParadas.Fields(2) = txtCadEscolaridade(2)
    rsParadas.Fields(3) = txtCadEscolaridade(1)
    rsParadas.Fields(4) = txtCadEscolaridade(3)
    If Check1.Value = 0 Then
        rsParadas.Fields(5) = "N"
    Else
        rsParadas.Fields(5) = "S"
    End If
    rsParadas.Update
'    cnBanco.CommitTrans
    rsParadas.Close
    Set rsParadas = Nothing
    AtualizaListview
    mobjMsg.Abrir "Os dados da Movimentação foram salvos com sucesso", Ok, informacao, "ZEUS"
    Unload Me
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
'    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
'    cnBanco.RollbackTrans
'    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DesbloqueiaControles
    For X = 0 To txtCadEscolaridade.Count - 1
        txtCadEscolaridade(X) = ""
    Next
    txtCadEscolaridade(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadEscolaridade(0).Text = Format(rsParadas.Fields(0), "000000")
    txtCadEscolaridade(1).Text = rsParadas.Fields(3)
    txtCadEscolaridade(2).Text = rsParadas.Fields(2)
    txtCadEscolaridade(3).Text = rsParadas.Fields(4)
    cboCadEscolaridade.Text = rsParadas.Fields(1)
    If rsParadas.Fields(5) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If cboCadEscolaridade.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.cboCadEscolaridade.Tag, Ok, critico, "Atenção"
        Me.cboCadEscolaridade.SetFocus
        Exit Function
    End If
    If txtCadEscolaridade(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadEscolaridade(1).Tag, Ok, critico, "Atenção"
        Me.txtCadEscolaridade(1).SetFocus
        Exit Function
    End If
    If txtCadEscolaridade(2).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadEscolaridade(0).Tag, Ok, critico, "Atenção"
        Me.txtCadEscolaridade(0).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadEscolaridade.Count - 1
        txtCadEscolaridade(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadEscolaridade.Count - 1
        txtCadEscolaridade(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirEscolaridade
    SqlGera = "Select top 1 * from tbParadas order by idparada Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsParadas.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadEscolaridade(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharEscolaridade
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

Private Sub AbrirEscolaridade()
On Error GoTo Err
    sqlParadas = "Select * from tbParadas Order by idparada"
    rsParadas.Open sqlParadas, cnBanco, adOpenKeyset, adLockOptimistic
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

Private Sub FecharEscolaridade()
    rsParadas.Close
    Set rsParadas = Nothing
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    sqlParadas = "Select * from tbParadas as a Where a.idparada= '" & Val(varGlobal) & "' order by idparada"
    rsParadas.Open sqlParadas, cnBanco, adOpenKeyset, adLockReadOnly
    If rsParadas.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Código da Movimentação não encontrado", Ok, critico, "Atenção"
    End If
    rsParadas.Close
    Set rsParadas = Nothing
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

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadEscolaridade(0), "000000"))
        ItemLst.SubItems(1) = cboCadEscolaridade.Text
        ItemLst.SubItems(2) = txtCadEscolaridade(2).Text
        ItemLst.SubItems(3) = txtCadEscolaridade(1).Text
        ItemLst.SubItems(4) = txtCadEscolaridade(3).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = cboCadEscolaridade.Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadEscolaridade(2).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadEscolaridade(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtCadEscolaridade(3).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub configControles()
    If vSal = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub txtCadEscolaridade_GotFocus(Index As Integer)
    mudaCorText txtCadEscolaridade(Index)
End Sub

Private Sub txtCadEscolaridade_LostFocus(Index As Integer)
    voltaCorText txtCadEscolaridade(Index)
End Sub
