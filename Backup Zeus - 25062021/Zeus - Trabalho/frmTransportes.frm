VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmTransportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Transportadoras"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmTransportes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6360
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Transportadora "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   3
         Left            =   3720
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   6015
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   5
         Left            =   6240
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   7
         Left            =   3480
         TabIndex        =   7
         Top             =   2280
         Width           =   2895
      End
      Begin VB.ComboBox cboCadEscolaridade 
         Height          =   315
         Index           =   0
         ItemData        =   "frmTransportes.frx":0CCA
         Left            =   6480
         List            =   "frmTransportes.frx":0D1F
         TabIndex        =   8
         Top             =   2280
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "frmTransportes.frx":0D8F
         TabIndex        =   21
         Top             =   2040
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "frmTransportes.frx":0DF3
         TabIndex        =   20
         Top             =   2040
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTransportes.frx":0E5F
         TabIndex        =   19
         Top             =   2040
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   6240
         OleObjectBlob   =   "frmTransportes.frx":0ECB
         TabIndex        =   18
         Top             =   1440
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTransportes.frx":0F31
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "frmTransportes.frx":0FA1
         TabIndex        =   16
         Top             =   840
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTransportes.frx":1025
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtCadEscolaridade 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Descrição da habilidade"
         ToolTipText     =   "Descrição da habilidade"
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox txtCadEscolaridade 
         Enabled         =   0   'False
         Height          =   315
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
         OleObjectBlob   =   "frmTransportes.frx":108D
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTransportes.frx":10F5
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   3120
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
      MICON           =   "frmTransportes.frx":1161
      PICN            =   "frmTransportes.frx":117D
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
      TabIndex        =   9
      Tag             =   "Salvar registro"
      ToolTipText     =   "Salvar registro"
      Top             =   3120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmTransportes.frx":1E57
      PICN            =   "frmTransportes.frx":1E73
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
Attribute VB_Name = "frmTransportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTransportadoras As New ADODB.Recordset
Private sqlTransportadoras As String
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados da Transportadora?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravarDados
'            gravaLog "Código esc.: " & txtCadEscolaridade(0), "Nome esc: " & txtCadEscolaridade(1), "Peso: " & txtCadEscolaridade(2)
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Transportadoras?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Unload Me
            Set frmTransportes = Nothing
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

Private Sub chamCad_Click(Index As Integer)

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
    'If Status = "novo" Then
    '    LimpaControles
    'ElseIf Status = "editar" Then
        ResultPesq
        DesbloqueiaControles
    'End If
    'configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub GravarDados()
On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarTransp As New ADODB.Recordset
    Dim sqlSalvarTransp As String
    Dim Y As Integer
    cnBanco.BeginTrans
   
    sqlSalvarTransp = "select * from tbTransportadoras where codtransp = '" & txtCadEscolaridade(0) & "'"
    rsSalvarTransp.Open sqlSalvarTransp, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarTransp.EOF Then rsSalvarTransp.AddNew
    rsSalvarTransp.Fields(0) = Val(txtCadEscolaridade(0))
    rsSalvarTransp.Fields(1) = txtCadEscolaridade(1)
    rsSalvarTransp.Fields(2) = txtCadEscolaridade(2)
    rsSalvarTransp.Fields(3) = txtCadEscolaridade(3)
    rsSalvarTransp.Fields(4) = txtCadEscolaridade(4)
    rsSalvarTransp.Fields(5) = txtCadEscolaridade(5)
    rsSalvarTransp.Fields(6) = txtCadEscolaridade(6)
    rsSalvarTransp.Fields(7) = txtCadEscolaridade(7)
    rsSalvarTransp.Fields(8) = cboCadEscolaridade(0)
    
    If Check1.Value = 0 Then
        rsSalvarTransp.Fields(9) = "N"
    Else
        rsSalvarTransp.Fields(9) = "S"
    End If
    rsSalvarTransp.Update
    
    cnBanco.CommitTrans
    
    rsSalvarTransp.Close
    Set rsSalvarTransp = Nothing
    AtualizaListview
    mobjMsg.Abrir "Os dados da Transportadora foram salvos com sucesso", Ok, informacao, "ZEUS"
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
    txtCadEscolaridade(0).Text = Format(rsTransportadoras.Fields(0), "000000")
    txtCadEscolaridade(1).Text = rsTransportadoras.Fields(1)
    If Not IsNull(rsTransportadoras.Fields(2)) Then txtCadEscolaridade(2).Text = rsTransportadoras.Fields(2) Else txtCadEscolaridade(2).Text = "-"
    If Not IsNull(rsTransportadoras.Fields(3)) Then txtCadEscolaridade(3).Text = rsTransportadoras.Fields(3) Else txtCadEscolaridade(3).Text = "-"
    If Not IsNull(rsTransportadoras.Fields(4)) Then txtCadEscolaridade(4).Text = rsTransportadoras.Fields(4) Else txtCadEscolaridade(4).Text = "-"
    If Not IsNull(rsTransportadoras.Fields(5)) Then txtCadEscolaridade(5).Text = rsTransportadoras.Fields(5) Else txtCadEscolaridade(5).Text = "-"
    If Not IsNull(rsTransportadoras.Fields(6)) Then txtCadEscolaridade(6).Text = rsTransportadoras.Fields(6) Else txtCadEscolaridade(6).Text = "-"
    If Not IsNull(rsTransportadoras.Fields(7)) Then txtCadEscolaridade(7).Text = rsTransportadoras.Fields(7) Else txtCadEscolaridade(7).Text = "-"
    If Not IsNull(rsTransportadoras.Fields(8)) Then cboCadEscolaridade(0).Text = rsTransportadoras.Fields(8) Else cboCadEscolaridade(0).Text = "-"
    'If rsTransportadoras.Fields(9) = "S" Then
    '    Check1.Value = 1
    'Else
    '    Check1.Value = 0
    'End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    Dim Y As Integer, X As Integer
    For X = 0 To 7
        If txtCadEscolaridade(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtCadEscolaridade(X).Tag, Ok, critico, "Atenção"
            Me.txtCadEscolaridade(X).SetFocus
            Exit Function
        End If
    Next
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
    SqlGera = "Select top 1 * from tbTransportadoras order by codtransp Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTransportadoras.RecordCount > 0 Then
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
    sqlTransportadoras = "Select * from tbTransportadoras Order by codtransp"
    rsTransportadoras.Open sqlTransportadoras, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharEscolaridade()
    rsTransportadoras.Close
    Set rsTransportadoras = Nothing
End Sub

Private Sub ResultPesq()
'    sqlTransportadoras = "Select * from tbTransportadoras Where tbTransportadoras.codtransp= '" & Val(varGlobal) & "' order by codtransp"
    sqlTransportadoras = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD,a.INATIVO from " & vBancoTotvs & ".dbo.ttra as a Where a.CODTRA= '" & varGlobal & "' order by CODTRA"
    rsTransportadoras.Open sqlTransportadoras, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTransportadoras.RecordCount > 0 Then
        CompoeControles
    'Else
    '    mobjMsg.Abrir "Código da Transportadora não encontrado", Ok, critico, "Atenção"
    End If
    rsTransportadoras.Close
    Set rsTransportadoras = Nothing
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
        ItemLst.SubItems(3) = txtCadEscolaridade(3).Text
        ItemLst.SubItems(4) = txtCadEscolaridade(4).Text
        ItemLst.SubItems(5) = txtCadEscolaridade(5).Text
        ItemLst.SubItems(6) = txtCadEscolaridade(6).Text
        ItemLst.SubItems(7) = txtCadEscolaridade(7).Text
        ItemLst.SubItems(8) = cboCadEscolaridade(0).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(9) = ""
            ItemLst.ListSubItems.Item(9).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(9) = ""
            ItemLst.ListSubItems.Item(9).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadEscolaridade(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadEscolaridade(2).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadEscolaridade(3).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtCadEscolaridade(4).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = txtCadEscolaridade(5).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = txtCadEscolaridade(6).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = txtCadEscolaridade(7).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = cboCadEscolaridade(0).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(9).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(9).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

'Private Sub configControles()
'    If vSal = "N" Then
'        cmdCadastro(0).UseGreyscale = True
'        cmdCadastro(0).DragMode = 1
'        cmdCadastro(0).SpecialEffect = cbEngraved
'    End If
'End Sub





