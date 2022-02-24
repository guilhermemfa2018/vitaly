VERSION 5.00
Begin VB.Form frmCargos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cargos"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmCargos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do cargo"
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8295
      Begin VB.Frame Frame2 
         Caption         =   "Descrição do cargo"
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   8055
         Begin VB.TextBox txtCadCargo 
            Height          =   1095
            Index           =   3
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Tag             =   "Breve descrição do cargo"
            ToolTipText     =   "Breve descrição do cargo"
            Top             =   240
            Width           =   7815
         End
      End
      Begin VB.TextBox txtCadCargo 
         Height          =   285
         Index           =   2
         Left            =   5880
         TabIndex        =   2
         Tag             =   "Número do CBO do cargo"
         ToolTipText     =   "Número do CBO do cargo"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtCadCargo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código  do cargo"
         ToolTipText     =   "Código  do cargo"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadCargo 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do cargo"
         ToolTipText     =   "Nome do cargo"
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "CBO nº:"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nome do cargo:"
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
      Top             =   2640
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
      MICON           =   "frmCargos.frx":0CCA
      PICN            =   "frmCargos.frx":0CE6
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
      Top             =   2640
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
      MICON           =   "frmCargos.frx":19C0
      PICN            =   "frmCargos.frx":19DC
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
Attribute VB_Name = "frmCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCargos As New ADODB.Recordset
Private SqlCargos As String
Private Status As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If MsgBox("Deseja salvar os dados do cargo?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravarDados
            gravaLog "ID: " & txtCadCargo(0), "Cargo: " & txtCadCargo(1), "CBO:" & txtCadCargo(2)
            'AtivaLD
            Unload Me
            Set frmCargos = Nothing
        End If
    Case 1
        If MsgBox("Deseja sair da tela de cadastro de cargos?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Unload Me
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
    Dim rsSalvarCargo As New ADODB.Recordset
    Dim SqlSalvarCargo As String
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarCargo = "select * from tbcargos where codcoligada = '" & vCodcoligada & "' and codcargo = '" & txtCadCargo(0) & "'"
    rsSalvarCargo.Open SqlSalvarCargo, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarCargo.EOF Then rsSalvarCargo.AddNew
    rsSalvarCargo.Fields(0) = Val(txtCadCargo(0))
    rsSalvarCargo.Fields(1) = txtCadCargo(2)
    rsSalvarCargo.Fields(2) = txtCadCargo(1)
    rsSalvarCargo.Fields(3) = txtCadCargo(3)
    If Check1.Value = 0 Then
        rsSalvarCargo.Fields(4) = "N"
    Else
        rsSalvarCargo.Fields(4) = "S"
    End If
    rsSalvarCargo.Fields(5) = vCodcoligada 'Codigo da coligada
    rsSalvarCargo.Update
'    X = rsMaterial.AbsolutePosition
    
    cnBanco.CommitTrans
    
    rsSalvarCargo.Close
    Set rsSalvarCargo = Nothing
    MsgBox "Os dados do cargo foram salvos com sucesso", vbInformation, "SGCH"
    AtualizaListview
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DesbloqueiaControles
    For X = 0 To txtCadCargo.Count - 1
        txtCadCargo(X) = ""
    Next
    txtCadCargo(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadCargo(0).Text = Format(rsCargos.Fields(0), "000000")
    txtCadCargo(1).Text = rsCargos.Fields(2)
    txtCadCargo(2).Text = rsCargos.Fields(1)
    If Not IsNull(rsCargos.Fields(3)) Then txtCadCargo(3).Text = rsCargos.Fields(3)
    If rsCargos.Fields(4) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadCargo(0).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadCargo(0).Tag, vbInformation, "Atenção"
        Me.txtCadCargo(0).SetFocus
        Exit Function
    End If
    If txtCadCargo(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadCargo(1).Tag, vbInformation, "Atenção"
        Me.txtCadCargo(1).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadCargo.Count - 1
        txtCadCargo(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadCargo.Count - 1
        txtCadCargo(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirCargo
    SqlGera = "Select top 1 * from tbCargos where codcoligada = '" & vCodcoligada & "' order by codcargo Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCargos.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadCargo(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharCargos
End Function

Private Sub AbrirCargo()
    SqlCargos = "Select * from tbCargos Where codcoligada = '" & vCodcoligada & "' Order by codcargo"
    rsCargos.Open SqlCargos, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharCargos()
    rsCargos.Close
    Set rsCargos = Nothing
End Sub

Private Sub ResultPesq()
    SqlCargos = "Select * from tbcargos Where codcoligada = '" & vCodcoligada & "' and tbcargos.codcargo= '" & Val(varGlobal) & "' order by codcargo"
    rsCargos.Open SqlCargos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCargos.RecordCount > 0 Then
        CompoeControles
    Else
        MsgBox "Cargo não encontrado"
    End If
    rsCargos.Close
    Set rsCargos = Nothing
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
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadCargo(0), "000000"))
        ItemLst.SubItems(1) = txtCadCargo(1).Text
        ItemLst.SubItems(2) = txtCadCargo(2).Text
        ItemLst.SubItems(3) = txtCadCargo(3).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadCargo(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadCargo(2).Text
        If txtCadCargo(3).Text <> "" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadCargo(3).Text
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
    MsgBox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub configControles()
    If vSal = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    End If
End Sub

