VERSION 5.00
Begin VB.Form frmAvaliacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avalia��es"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmAvaliacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da avalia��o"
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtCadAvaliacao 
         Height          =   1215
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1800
         Width           =   8055
      End
      Begin VB.ComboBox cboCadAvaliacao 
         Height          =   315
         ItemData        =   "frmAvaliacoes.frx":0CCA
         Left            =   120
         List            =   "frmAvaliacoes.frx":0CD7
         TabIndex        =   3
         Tag             =   "Tipo da avalia��o"
         Text            =   "AE (Avalia��o de Efic�cia do Treinamento)"
         ToolTipText     =   "Tipo da avalia��o"
         Top             =   1080
         Width           =   7335
      End
      Begin VB.TextBox txtCadAvaliacao 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "C�digo  do treinamento"
         ToolTipText     =   "C�digo  do treinamento"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadAvaliacao 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Conte�do da avalia��o"
         ToolTipText     =   "Conte�do da avalia��o"
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtCadAvaliacao 
         Height          =   285
         Index           =   2
         Left            =   7560
         TabIndex        =   2
         Tag             =   "Peso da avalia��o"
         ToolTipText     =   "Peso da avalia��o"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Descri��o:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Peso:"
         Height          =   255
         Left            =   7560
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Avalia��o:"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   3240
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
      MICON           =   "frmAvaliacoes.frx":0D43
      PICN            =   "frmAvaliacoes.frx":0D5F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   3240
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
      MICON           =   "frmAvaliacoes.frx":1A39
      PICN            =   "frmAvaliacoes.frx":1A55
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
Attribute VB_Name = "frmAvaliacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsAvaliacoes As New ADODB.Recordset
Private sqlAvaliacoes As String
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If MsgBox("Deseja salvar os dados da Avalia��o?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravarDados
            gravaLog "C�digo AV: " & txtCadAvaliacao(0) & "Nome AV: " & txtCadAvaliacao(1), "Peso: " & txtCadAvaliacao(2), "Tipo:" & cboCadAvaliacao
            Unload Me
            Set frmAvaliacoes = Nothing
        End If
    Case 1
        If MsgBox("Deseja sair da tela de cadastro de Avalia��es?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Unload Me
            Set frmAvaliacoes = Nothing
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
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarAvaliacoes As New ADODB.Recordset
    Dim SqlSalvarAvaliacoes As String
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarAvaliacoes = "select * from tbAvaliacao where codcoligada = '" & vCodcoligada & "' and codAvaliacao = '" & txtCadAvaliacao(0) & "'"
    rsSalvarAvaliacoes.Open SqlSalvarAvaliacoes, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarAvaliacoes.EOF Then rsSalvarAvaliacoes.AddNew
    rsSalvarAvaliacoes.Fields(0) = Val(txtCadAvaliacao(0)) 'C�digo da avalia��o
    rsSalvarAvaliacoes.Fields(1) = txtCadAvaliacao(1) 'Conte�do da avalia��o
    rsSalvarAvaliacoes.Fields(3) = txtCadAvaliacao(2) 'Peso da avalia��o
    If cboCadAvaliacao = "AE (Avalia��o de Efic�cia do Treinamento)" Then
        rsSalvarAvaliacoes.Fields(2) = "AE" 'Tipo da avalia��o
    ElseIf cboCadAvaliacao = "AT (Avalia��o do Treinamento)" Then
        rsSalvarAvaliacoes.Fields(2) = "AT" 'Tipo da avalia��o
    Else
        rsSalvarAvaliacoes.Fields(2) = "AD" 'Tipo da avalia��o
    End If
    If Check1.Value = 0 Then
        rsSalvarAvaliacoes.Fields(4) = "N" 'Nao ativo
    Else
        rsSalvarAvaliacoes.Fields(4) = "S" 'Ativo
    End If
    rsSalvarAvaliacoes.Fields(5) = txtCadAvaliacao(3) 'Descri��o
    rsSalvarAvaliacoes.Fields(6) = vCodcoligada ' Codigo da coligada
    rsSalvarAvaliacoes.Update
    
    cnBanco.CommitTrans
    
    rsSalvarAvaliacoes.Close
    Set rsSalvarAvaliacoes = Nothing
    AtualizaListview
    MsgBox "Os dados da Avalia��o foram salvos com sucesso", vbInformation, "SGCH"
    'Unload Me
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alter��es nos registros ser�o desfeitas!", vbInformation, "Aten��o"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DesbloqueiaControles
    For X = 0 To txtCadAvaliacao.Count - 1
        txtCadAvaliacao(X) = ""
    Next
    cboCadAvaliacao.Text = "AE (Avalia��o de Efic�cia do Treinamento)"
    txtCadAvaliacao(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadAvaliacao(0).Text = Format(rsAvaliacoes.Fields(0), "000000")
    txtCadAvaliacao(1).Text = rsAvaliacoes.Fields(1)
    txtCadAvaliacao(2).Text = rsAvaliacoes.Fields(3)
    If rsAvaliacoes.Fields(2) = "AE" Then
        cboCadAvaliacao.Text = "AE (Avalia��o de Efic�cia do Treinamento)"
    ElseIf rsAvaliacoes.Fields(2) = "AT" Then
        cboCadAvaliacao.Text = "AT (Avalia��o do Treinamento)"
    Else
        cboCadAvaliacao.Text = "AD (Avalia��o de Desempenho)"
    End If
    If rsAvaliacoes.Fields(4) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If Not IsNull(rsAvaliacoes.Fields(5)) Then txtCadAvaliacao(3).Text = rsAvaliacoes.Fields(5)
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadAvaliacao(0).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadAvaliacao(0).Tag, vbInformation, "Aten��o"
        Me.txtCadAvaliacao(0).SetFocus
        Exit Function
    End If
    If txtCadAvaliacao(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadAvaliacao(1).Tag, vbInformation, "Aten��o"
        Me.txtCadAvaliacao(1).SetFocus
        Exit Function
    End If
    If txtCadAvaliacao(2).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadAvaliacao(2).Tag, vbInformation, "Aten��o"
        Me.txtCadAvaliacao(2).SetFocus
        Exit Function
    End If
    
    If cboCadAvaliacao.Text = "" Then
        MsgBox "Favor informar o campo " & Me.cboCadAvaliacao.Tag, vbInformation, "Aten��o"
        Me.cboCadAvaliacao.SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadAvaliacao.Count - 1
        txtCadAvaliacao(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadAvaliacao.Count - 1
        txtCadAvaliacao(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    AbrirAvaliacoes
    SqlGera = "Select top 1 * from tbAvaliacao where codcoligada = '" & vCodcoligada & "' order by codAvaliacao Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAvaliacoes.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadAvaliacao(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharAvaliacoes
End Function

Private Sub AbrirAvaliacoes()
    sqlAvaliacoes = "Select * from tbAvaliacao where codcoligada ='" & vCodcoligada & "' Order by codAvaliacao"
    rsAvaliacoes.Open sqlAvaliacoes, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharAvaliacoes()
    rsAvaliacoes.Close
    Set rsAvaliacoes = Nothing
End Sub

Private Sub ResultPesq()
    sqlAvaliacoes = "Select * from tbAvaliacao Where codcoligada = '" & vCodcoligada & "' and tbAvaliacao.codAvaliacao= '" & Val(varGlobal) & "' order by codAvaliacao"
    rsAvaliacoes.Open sqlAvaliacoes, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAvaliacoes.RecordCount > 0 Then
        CompoeControles
    Else
        MsgBox "Avalia��o n�o encontrado"
    End If
    rsAvaliacoes.Close
    Set rsAvaliacoes = Nothing
End Sub

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadAvaliacao(0), "000000"))
        ItemLst.SubItems(1) = txtCadAvaliacao(1).Text
        If cboCadAvaliacao.Text = "AE (Avalia��o de Efic�cia do Treinamento)" Then
            ItemLst.SubItems(2) = "AE"
        ElseIf cboCadAvaliacao.Text = "AT (Avalia��o do Treinamento)" Then
            ItemLst.SubItems(2) = "AT"
        Else
            ItemLst.SubItems(2) = "AD"
        End If
        ItemLst.SubItems(3) = txtCadAvaliacao(2).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadAvaliacao(1).Text
        If cboCadAvaliacao.Text = "AE (Avalia��o de Efic�cia do Treinamento)" Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = "AE"
        ElseIf cboCadAvaliacao.Text = "AT (Avalia��o do Treinamento)" Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = "AT"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = "AD"
        End If
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadAvaliacao(2).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    MsgBox "N�o foi poss�vel realizar as altera��es", vbInformation, "Aten��o"
    Exit Sub
End Sub

Private Sub configControles()
    If vSal = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    End If
End Sub


