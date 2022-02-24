VERSION 5.00
Begin VB.Form frmEscolaridade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolaridade"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmEscolaridade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da escolaridade"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtCadEscolaridade 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "C�digo  da habilidade"
         ToolTipText     =   "C�digo  da habilidade"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Descri��o da habilidade"
         ToolTipText     =   "Descri��o da habilidade"
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   285
         Index           =   2
         Left            =   7560
         TabIndex        =   2
         Tag             =   "Peso da habilidade"
         ToolTipText     =   "Peso da habilidade"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Peso:"
         Height          =   255
         Left            =   7560
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Escolaridade:"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   5
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
      MICON           =   "frmEscolaridade.frx":0CCA
      PICN            =   "frmEscolaridade.frx":0CE6
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
      TabIndex        =   4
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
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
      MICON           =   "frmEscolaridade.frx":19C0
      PICN            =   "frmEscolaridade.frx":19DC
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
Attribute VB_Name = "frmEscolaridade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsEscolaridade As New ADODB.Recordset
Private sqlEscolaridade As String
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If MsgBox("Deseja salvar os dados da Escolaridade?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravarDados
            gravaLog "C�digo esc.: " & txtCadEscolaridade(0), "Nome esc: " & txtCadEscolaridade(1), "Peso: " & txtCadEscolaridade(2)
        End If
    Case 1
        If MsgBox("Deseja sair da tela de cadastro de Escolaridade?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Unload Me
            Set frmEscolaridade = Nothing
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
On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarEscolaridade As New ADODB.Recordset
    Dim SqlSalvarEscolaridade As String
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarEscolaridade = "select * from tbEscolaridade where codcoligada = '" & vCodcoligada & "' and codEscolaridade = '" & txtCadEscolaridade(0) & "'"
    rsSalvarEscolaridade.Open SqlSalvarEscolaridade, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarEscolaridade.EOF Then rsSalvarEscolaridade.AddNew
    rsSalvarEscolaridade.Fields(0) = Val(txtCadEscolaridade(0))
    rsSalvarEscolaridade.Fields(1) = txtCadEscolaridade(1)
    rsSalvarEscolaridade.Fields(2) = txtCadEscolaridade(2)
    If Check1.Value = 0 Then
        rsSalvarEscolaridade.Fields(3) = "N"
    Else
        rsSalvarEscolaridade.Fields(3) = "S"
    End If
    rsSalvarEscolaridade.Fields(4) = vCodcoligada 'Codigo da coligada
    rsSalvarEscolaridade.Update
    
    cnBanco.CommitTrans
    
    rsSalvarEscolaridade.Close
    Set rsSalvarEscolaridade = Nothing
    AtualizaListview
    MsgBox "Os dados do Escolaridade foram salvos com sucesso", vbInformation, "SGCH"
    Unload Me
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alter��es nos registros ser�o desfeitas!", vbInformation, "Aten��o"
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
    txtCadEscolaridade(0).Text = Format(rsEscolaridade.Fields(0), "000000")
    txtCadEscolaridade(1).Text = rsEscolaridade.Fields(1)
    txtCadEscolaridade(2).Text = rsEscolaridade.Fields(2)
    If rsEscolaridade.Fields(3) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadEscolaridade(0).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadEscolaridade(0).Tag, vbInformation, "Aten��o"
        Me.txtCadEscolaridade(0).SetFocus
        Exit Function
    End If
    If txtCadEscolaridade(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadEscolaridade(1).Tag, vbInformation, "Aten��o"
        Me.txtCadEscolaridade(1).SetFocus
        Exit Function
    End If
    If txtCadEscolaridade(2).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadEscolaridade(2).Tag, vbInformation, "Aten��o"
        Me.txtCadEscolaridade(2).SetFocus
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
    SqlGera = "Select top 1 * from tbEscolaridade where codcoligada = '" & vCodcoligada & "' order by codEscolaridade Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsEscolaridade.RecordCount > 0 Then
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
    sqlEscolaridade = "Select * from tbEscolaridade where codcoligada = '" & vCodcoligada & "' Order by codEscolaridade"
    rsEscolaridade.Open sqlEscolaridade, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharEscolaridade()
    rsEscolaridade.Close
    Set rsEscolaridade = Nothing
End Sub

Private Sub ResultPesq()
    sqlEscolaridade = "Select * from tbEscolaridade Where tbEscolaridade.codcoligada = '" & vCodcoligada & "' and tbEscolaridade.codEscolaridade= '" & Val(varGlobal) & "' order by codEscolaridade"
    rsEscolaridade.Open sqlEscolaridade, cnBanco, adOpenKeyset, adLockReadOnly
    If rsEscolaridade.RecordCount > 0 Then
        CompoeControles
    Else
        MsgBox "Escolaridade n�o encontrado"
    End If
    rsEscolaridade.Close
    Set rsEscolaridade = Nothing
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
        ItemLst.SubItems(2) = txtCadEscolaridade(2).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadEscolaridade(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadEscolaridade(2).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "OK"
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

