VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTipoMat 
   Caption         =   "Cadastro de tipos de material"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "frmTipoMat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do tipo de material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Descrição da habilidade"
         ToolTipText     =   "Descrição da habilidade"
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtCadEscolaridade 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código  da habilidade"
         ToolTipText     =   "Código  da habilidade"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmTipoMat.frx":0CCA
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTipoMat.frx":0D32
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   1200
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
      MICON           =   "frmTipoMat.frx":0D9E
      PICN            =   "frmTipoMat.frx":0DBA
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
      TabIndex        =   2
      Tag             =   "Salvar registro"
      ToolTipText     =   "Salvar registro"
      Top             =   1200
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
      MICON           =   "frmTipoMat.frx":1A94
      PICN            =   "frmTipoMat.frx":1AB0
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
Attribute VB_Name = "frmTipoMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTipoMat As New ADODB.Recordset
Private sqlTipoMat As String
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados do Tipo de Material?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravarDados
'            gravaLog "Código esc.: " & txtCadEscolaridade(0), "Nome esc: " & txtCadEscolaridade(1), "Peso: " & txtCadEscolaridade(2)
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Tipo de Material?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Unload Me
            Set frmTipoMat = Nothing
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
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarTipoMat As New ADODB.Recordset
    Dim sqlSalvarTipoMat As String
    Dim Y As Integer
    cnBanco.BeginTrans
   
    sqlSalvarTipoMat = "select * from tbTipoMat where codigo = '" & txtCadEscolaridade(0) & "'"
    rsSalvarTipoMat.Open sqlSalvarTipoMat, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarTipoMat.EOF Then rsSalvarTipoMat.AddNew
    rsSalvarTipoMat.Fields(0) = Val(txtCadEscolaridade(0))
    rsSalvarTipoMat.Fields(1) = txtCadEscolaridade(1)
    If Check1.Value = 0 Then
        rsSalvarTipoMat.Fields(2) = "N"
    Else
        rsSalvarTipoMat.Fields(2) = "S"
    End If
    rsSalvarTipoMat.Update
    
    cnBanco.CommitTrans
    
    rsSalvarTipoMat.Close
    Set rsSalvarTipoMat = Nothing
    AtualizaListview
    mobjMsg.Abrir "Os dados do Tipo de Material foram salvos com sucesso", Ok, informacao, "ZEUS"
    Unload Me
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
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
    txtCadEscolaridade(0).Text = Format(rsTipoMat.Fields(0), "000000")
    txtCadEscolaridade(1).Text = rsTipoMat.Fields(1)
    If rsTipoMat.Fields(2) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadEscolaridade(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadEscolaridade(0).Tag, Ok, critico, "Atenção"
        Me.txtCadEscolaridade(0).SetFocus
        Exit Function
    End If
    If txtCadEscolaridade(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadEscolaridade(1).Tag, Ok, critico, "Atenção"
        Me.txtCadEscolaridade(1).SetFocus
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
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirEscolaridade
    SqlGera = "Select top 1 * from tbTipoMat order by codigo Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTipoMat.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadEscolaridade(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharEscolaridade
End Function

Private Sub AbrirEscolaridade()
    sqlTipoMat = "Select * from tbTipoMat Order by codigo"
    rsTipoMat.Open sqlTipoMat, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharEscolaridade()
    rsTipoMat.Close
    Set rsTipoMat = Nothing
End Sub

Private Sub ResultPesq()
    sqlTipoMat = "Select * from tbTipoMat Where tbTipoMat.codigo= '" & Val(varGlobal) & "' order by codigo"
    rsTipoMat.Open sqlTipoMat, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTipoMat.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Código do Material não encontrado", Ok, critico, "Atenção"
    End If
    rsTipoMat.Close
    Set rsTipoMat = Nothing
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
        ItemLst.SubItems(1) = txtCadEscolaridade(1).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadEscolaridade(1).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ReportIcon = "OK"
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




