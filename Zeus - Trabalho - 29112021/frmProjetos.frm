VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmProjetos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Projetos"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjetos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Projetos"
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   2400
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.CommandButton cmdcadastro 
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
      Index           =   11
      Left            =   120
      Picture         =   "frmProjetos.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "Sair"
      Top             =   8400
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Projeto "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton cmdcadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdcadastro 
         Height          =   615
         Index           =   7
         Left            =   1320
         Picture         =   "frmProjetos.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "Excluir"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdcadastro 
         Height          =   615
         Index           =   6
         Left            =   720
         Picture         =   "frmProjetos.frx":265E
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "Editar"
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdcadastro 
         Height          =   615
         Index           =   5
         Left            =   120
         Picture         =   "frmProjetos.frx":3328
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "Inserir"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtCadastro 
         Height          =   1125
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "Observa��o"
         Top             =   1920
         Width           =   10215
      End
      Begin VB.TextBox txtCadastro 
         Height          =   345
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Tag             =   "Descri��o"
         Top             =   1200
         Width           =   10215
      End
      Begin VB.TextBox txtCadastro 
         Height          =   345
         Index           =   5
         Left            =   9000
         TabIndex        =   3
         Tag             =   "N� Ordem de Compra do cliente"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
         Height          =   345
         Index           =   2
         Left            =   5520
         TabIndex        =   2
         Tag             =   "Projeto n�"
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtCadastro 
         Height          =   345
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         Tag             =   "FCE"
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1800
         TabIndex        =   8
         Tag             =   "Data de cadastro"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   162398209
         CurrentDate     =   40449
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "C�digo Identificador"
         ToolTipText     =   "C�digo"
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProjetos.frx":3FF2
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProjetos.frx":4060
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   9000
         OleObjectBlob   =   "frmProjetos.frx":40CC
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "frmProjetos.frx":4130
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "frmProjetos.frx":419E
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmProjetos.frx":4204
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProjetos.frx":4266
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7223
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
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
      End
   End
End
Attribute VB_Name = "frmProjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProjetos As New ADODB.Recordset
Private SqlProjetos As String
Private rsLocal As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridFCE
        txtcadastro(1).SetFocus
    Case 5
        IncluiTreeview
        txtcadastro(1).SetFocus
    Case 6
        'mskCadastro(2).PromptInclude = False
        'mskCadastro(2) = ""
        'mskCadastro(2).PromptInclude = True
        txtcadastro(2) = ""
        AlteraTreeview
    Case 7
        DeletaTreeview
        CompoeTreeview
    Case 10
        'mskCadastro_GotFocus (1)
        'ChamaGridGrupo
        'CarregaGrupo
    Case 11
        Unload Me
    Case 12
        'Bot_Salvar
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    frmProjetos.Left = 2710
    frmProjetos.Top = 0
    txtcadastro(0).Text = Format(GeraCodigo, "000000")
    CompoeTreeview
    carregarIconBotao
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte T�cnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdCadastro(5), 5, 46 'Inserir
    carregaImagemBotao cmdCadastro(6), 6, 32 'Editar
    carregaImagemBotao cmdCadastro(7), 7, 33 'Excluir
    carregaImagemBotao cmdCadastro(11), 11, 34 'Sair
End Sub

Private Sub AbrirProjeto()
On Error GoTo Err
    SqlProjetos = "Select * from tbProjetos Order by codprojeto"
    rsProjetos.Open SqlProjetos, cnBanco, adOpenKeyset, adLockReadOnly
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
    FecharProjeto
End Sub

Private Sub FecharProjeto()
    rsProjetos.Close
    Set rsProjetos = Nothing
End Sub

Private Sub LimpaControles()
    txtcadastro(0) = ""
    txtcadastro(2) = ""
    txtcadastro(3) = ""
    txtcadastro(4) = ""
    txtcadastro(5) = ""
'    For X = 0 To txtcadastro.Count - 1
'        txtcadastro(X) = ""
'    Next
    txtcadastro(0).Text = Format(GeraCodigo, "000000")
End Sub

Private Sub CompoeTreeview()
    Dim rsTree As New ADODB.Recordset
    Dim SqlTree
    Dim no As Node
    Dim x As Integer, y As Integer, ContaNo As Integer
    Dim FormataProj As String, FormataDesc As String
    SqlTree = "Select * from tbprojetos Order by fce"
    rsTree.Open SqlTree, cnBanco, adOpenKeyset, adLockOptimistic
    
    TreeView1.Nodes.Clear
    For x = 1 To rsTree.RecordCount
        Set no = TreeView1.Nodes.Add(, , "no" & x, "FCE:" & Format(rsTree.Fields(1), "000000"))
        ContaNo = ContaNo + 1
        'A linha abaixo server para expandir o NO
        'TreeView1.Nodes(ContaNo).Expanded = True
        y = rsTree.Fields(1)
        While y = rsTree.Fields(1)
            FormataProj = rsTree.Fields(2)
            FormataDesc = rsTree.Fields(3)
            While Len(FormataProj) < 20
                FormataProj = FormataProj + " "
            Wend
            While Len(FormataDesc) < 20
                FormataDesc = FormataDesc + " "
            Wend
            TreeView1.Nodes.Add "no" & x, tvwChild, , "ID:" & Format(rsTree.Fields(0), "000000") & "- PROJETO:" & FormataProj & "-DESC:" & FormataDesc & "- DATA:" & CStr(rsTree.Fields(4)) & "- OBS:" & Mid$(rsTree.Fields(5), 1, 20)
            rsTree.MoveNext
            ContaNo = ContaNo + 1
            'A linha abaixo server para expandir o NO
            'TreeView1.Nodes(ContaNo).Expanded = True
            If rsTree.EOF Then Exit For
        Wend
    Next
    rsTree.Close
    Set rsTree = Nothing
End Sub

Private Sub IncluiTreeview()
On Error GoTo Err
    Dim rsItem As New ADODB.Recordset
    Dim SqlItem As String
    If ValidaCampo = False Then Exit Sub
    SqlItem = "Select * from tbProjetos where tbProjetos.codprojeto = '" & Val(txtcadastro(0)) & "'"
    rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
    If rsItem.RecordCount = 0 Then
        rsItem.AddNew
        rsItem.Fields(0) = Val(txtcadastro(0))
        'rsItem.Fields(1) = Val(txtCadastro(1))
        'rsItem.Fields(2) = txtCadastro(2)
        'rsItem.Fields(3) = txtCadastro(3)
        'rsItem.Fields(4) = DTPicker1
        'rsItem.Fields(5) = txtCadastro(4)
        'rsItem.Fields(6) = txtCadastro(5)
    End If
    rsItem.Fields(1) = Val(txtcadastro(1))
    rsItem.Fields(2) = txtcadastro(2)
    rsItem.Fields(3) = txtcadastro(3)
    rsItem.Fields(4) = DTPicker1
    rsItem.Fields(5) = txtcadastro(4)
    rsItem.Fields(6) = txtcadastro(5)
    rsItem.Update
    Set rsItem = Nothing
    CompoeTreeview
    LimpaControles
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

Private Sub AlteraTreeview()
On Error GoTo Err
    Dim llng_Contador As Long
    
    Dim rsItem As New ADODB.Recordset
    Dim SqlItem As String
    
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") <> 0 Then
                'MsgBox "Subitem"
                LimpaControles
                txtcadastro(0) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 4, 6)
                SqlItem = "Select * from tbProjetos where tbProjetos.codprojeto = '" & Val(txtcadastro(0)) & "'"
                rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
                txtcadastro(1) = Format(rsItem.Fields(1), "000000")
                txtcadastro(2) = rsItem.Fields(2)
                txtcadastro(3) = rsItem.Fields(3)
                txtcadastro(4) = rsItem.Fields(5)
                txtcadastro(5) = rsItem.Fields(6)
                DTPicker1 = rsItem.Fields(4)
                rsItem.Close
                Set rsItem = Nothing
            Else
                'MsgBox "Grupo"
                LimpaControles
                AbrirProjeto
                txtcadastro(0).Text = Format(GeraCodigo, "000000")
                FecharProjeto
                txtcadastro(1) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 5, 6)
            End If
        End If
    Next
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

Private Sub DeletaTreeview()
On Error GoTo Err
    Dim llng_Contador As Long
    Dim rsItem As New ADODB.Recordset
    Dim SqlItem As String
    
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If Msgbox("Confirma Exclus�o", vbQuestion + vbYesNo, "ZEUS") = vbYes Then
                SqlItem = "Delete from tbProjetos where tbProjetos.codprojeto =" & " '" & Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 4, 6)) & "'"
                rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
            End If
        End If
    Next
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

Private Function ValidaCampo()
    ValidaCampo = False
    For x = 0 To 3
        If txtcadastro(x).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo: " & Me.txtcadastro(x).Tag, Ok, critico, "Aten��o"
            Me.txtcadastro(x).SetFocus
            Exit Function
        End If
    Next
    ValidaCampo = True
End Function

Private Sub CarregaFCE()
On Error GoTo Err
    Dim x As Integer
    Dim rsFCE As New ADODB.Recordset
    SqlM = "Select * from tbfce where tbfce.fce = '" & Val(txtcadastro(1)) & "' and tbfce.status = 0 order by fce"
    rsFCE.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsFCE.EOF Then rsFCE.MoveFirst
    If rsFCE.EOF Then
        txtcadastro(1).Text = Format(txtcadastro(1), "000000") & ""
        mobjMsg.Abrir "FCE n�o cadastrada", Ok, critico, "Aten��o"
        txtcadastro(1).SetFocus
    Else
        txtcadastro(1).Text = Format(rsFCE.Fields(0), "000000") & ""
        txtcadastro(2).SetFocus
    End If
    rsFCE.Close
    Set rsFCE = Nothing
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

Private Sub ChamaGridFCE()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbFCE where status = 0 order by FCE"
    procnom = "fce"
    campo = 0
    Campo1 = 3
    Load F
    F.Caption = "Pesquisa de FCEs"
    Pesquisa = frmProjetos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "fce=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(1).Text = Format(rsLocal.Fields(0), "000000")
            txtcadastro(2).SetFocus
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
        Resume Next
    End If
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirProjeto
    SqlGera = "Select top 1 * from tbProjetos order by codprojeto Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProjetos.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtcadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharProjeto
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

Private Sub TreeView1_DblClick()
    AlteraTreeview
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            CarregaFCE
        End If
    End If
End Sub

Private Sub txtCadastro_LostFocus(Index As Integer)
'    If Index = 1 Then
'        CarregaFCE
'    End If
End Sub
