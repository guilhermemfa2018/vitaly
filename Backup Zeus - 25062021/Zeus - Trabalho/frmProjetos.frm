VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjetos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Projetos"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmProjetos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Projetos"
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Projeto "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10455
      Begin VB.TextBox txtCadastro 
         Height          =   1125
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   1680
         Width           =   10215
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Tag             =   "Descrição"
         ToolTipText     =   "Descrição"
         Top             =   1080
         Width           =   10215
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   5
         Left            =   9000
         TabIndex        =   4
         ToolTipText     =   "Nº Ordem de Compra do cliente"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   3
         Tag             =   "Projeto nº"
         ToolTipText     =   "Projeto nº"
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   1
         Tag             =   "FCE"
         ToolTipText     =   "FCE"
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   54525953
         CurrentDate     =   40449
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código"
         ToolTipText     =   "Código"
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProjetos.frx":0CCA
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProjetos.frx":0D3E
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   9000
         OleObjectBlob   =   "frmProjetos.frx":0DB0
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "frmProjetos.frx":0E1A
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "frmProjetos.frx":0E8E
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmProjetos.frx":0EFA
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProjetos.frx":0F62
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin ZEUS.chameleonButton cmdcadastro 
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   4
         TX              =   "..."
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
         MICON           =   "frmProjetos.frx":0FCE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6165
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin ZEUS.chameleonButton cmdcadastro 
         Height          =   615
         Index           =   7
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
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
         MICON           =   "frmProjetos.frx":0FEA
         PICN            =   "frmProjetos.frx":1006
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdcadastro 
         Height          =   615
         Index           =   6
         Left            =   720
         TabIndex        =   8
         Top             =   3000
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
         MICON           =   "frmProjetos.frx":1CE0
         PICN            =   "frmProjetos.frx":1CFC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdcadastro 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   3000
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
         MICON           =   "frmProjetos.frx":29D6
         PICN            =   "frmProjetos.frx":29F2
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
   Begin ZEUS.chameleonButton cmdcadastro 
      Height          =   615
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   7560
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
      MICON           =   "frmProjetos.frx":36CC
      PICN            =   "frmProjetos.frx":36E8
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
        txtCadastro(1).SetFocus
    Case 5
        IncluiTreeview
        txtCadastro(1).SetFocus
    Case 6
        'mskCadastro(2).PromptInclude = False
        'mskCadastro(2) = ""
        'mskCadastro(2).PromptInclude = True
        txtCadastro(2) = ""
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
    txtCadastro(0).Text = Format(GeraCodigo, "000000")
    CompoeTreeview
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub AbrirProjeto()
    On Error GoTo Err
    SqlProjetos = "Select * from tbProjetos Order by codprojeto"
    rsProjetos.Open SqlProjetos, cnBanco, adOpenKeyset, adLockReadOnly
    Exit Sub
Err:
    FecharProjeto
End Sub

Private Sub FecharProjeto()
    rsProjetos.Close
    Set rsProjetos = Nothing
End Sub

Private Sub LimpaControles()
    txtCadastro(0) = ""
    txtCadastro(2) = ""
    txtCadastro(3) = ""
    txtCadastro(4) = ""
    txtCadastro(5) = ""
'    For X = 0 To txtcadastro.Count - 1
'        txtcadastro(X) = ""
'    Next
    txtCadastro(0).Text = Format(GeraCodigo, "000000")
End Sub

Private Sub CompoeTreeview()
    Dim rsTree As New ADODB.Recordset
    Dim SqlTree
    Dim no As Node
    Dim X As Integer, Y As Integer, ContaNo As Integer
    Dim FormataProj As String, FormataDesc As String
    SqlTree = "Select * from tbprojetos Order by fce"
    rsTree.Open SqlTree, cnBanco, adOpenKeyset, adLockOptimistic
    
    TreeView1.Nodes.Clear
    For X = 1 To rsTree.RecordCount
        Set no = TreeView1.Nodes.Add(, , "no" & X, "FCE:" & Format(rsTree.Fields(1), "000000"))
        ContaNo = ContaNo + 1
        'A linha abaixo server para expandir o NO
        'TreeView1.Nodes(ContaNo).Expanded = True
        Y = rsTree.Fields(1)
        While Y = rsTree.Fields(1)
            FormataProj = rsTree.Fields(2)
            FormataDesc = rsTree.Fields(3)
            While Len(FormataProj) < 20
                FormataProj = FormataProj + " "
            Wend
            While Len(FormataDesc) < 20
                FormataDesc = FormataDesc + " "
            Wend
            TreeView1.Nodes.Add "no" & X, tvwChild, , "ID:" & Format(rsTree.Fields(0), "000000") & "- PROJETO:" & FormataProj & "-DESC:" & FormataDesc & "- DATA:" & CStr(rsTree.Fields(4)) & "- OBS:" & Mid$(rsTree.Fields(5), 1, 20)
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
    Dim rsItem As New ADODB.Recordset
    Dim SqlItem As String
    If ValidaCampo = False Then Exit Sub
    SqlItem = "Select * from tbProjetos where tbProjetos.codprojeto = '" & Val(txtCadastro(0)) & "'"
    rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
    If rsItem.RecordCount = 0 Then
        rsItem.AddNew
        rsItem.Fields(0) = Val(txtCadastro(0))
        'rsItem.Fields(1) = Val(txtCadastro(1))
        'rsItem.Fields(2) = txtCadastro(2)
        'rsItem.Fields(3) = txtCadastro(3)
        'rsItem.Fields(4) = DTPicker1
        'rsItem.Fields(5) = txtCadastro(4)
        'rsItem.Fields(6) = txtCadastro(5)
    End If
    rsItem.Fields(1) = Val(txtCadastro(1))
    rsItem.Fields(2) = txtCadastro(2)
    rsItem.Fields(3) = txtCadastro(3)
    rsItem.Fields(4) = DTPicker1
    rsItem.Fields(5) = txtCadastro(4)
    rsItem.Fields(6) = txtCadastro(5)
    rsItem.Update
    Set rsItem = Nothing
    CompoeTreeview
    LimpaControles
End Sub

Private Sub AlteraTreeview()
    Dim llng_Contador As Long
    
    Dim rsItem As New ADODB.Recordset
    Dim SqlItem As String
    
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") <> 0 Then
                'MsgBox "Subitem"
                LimpaControles
                txtCadastro(0) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 4, 6)
                SqlItem = "Select * from tbProjetos where tbProjetos.codprojeto = '" & Val(txtCadastro(0)) & "'"
                rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
                txtCadastro(1) = Format(rsItem.Fields(1), "000000")
                txtCadastro(2) = rsItem.Fields(2)
                txtCadastro(3) = rsItem.Fields(3)
                txtCadastro(4) = rsItem.Fields(5)
                txtCadastro(5) = rsItem.Fields(6)
                DTPicker1 = rsItem.Fields(4)
                rsItem.Close
                Set rsItem = Nothing
            Else
                'MsgBox "Grupo"
                LimpaControles
                AbrirProjeto
                txtCadastro(0).Text = Format(GeraCodigo, "000000")
                FecharProjeto
                txtCadastro(1) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 5, 6)
            End If
        End If
    Next
End Sub

Private Sub DeletaTreeview()
    Dim llng_Contador As Long
    Dim rsItem As New ADODB.Recordset
    Dim SqlItem As String
    
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If Msgbox("Confirma Exclusão", vbQuestion + vbYesNo, "ZEUS") = vbYes Then
                SqlItem = "Delete from tbProjetos where tbProjetos.codprojeto =" & " '" & Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 4, 6)) & "'"
                rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
            End If
        End If
    Next
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    For X = 0 To 3
        If txtCadastro(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo: " & Me.txtCadastro(X).Tag, Ok, critico, "Atenção"
            Me.txtCadastro(X).SetFocus
            Exit Function
        End If
    Next
    ValidaCampo = True
End Function

Private Sub CarregaFCE()
    Dim X As Integer
    Dim rsFCE As New ADODB.Recordset
    SqlM = "Select * from tbfce where tbfce.fce = '" & Val(txtCadastro(1)) & "' and tbfce.status = 0 order by fce"
    rsFCE.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsFCE.EOF Then rsFCE.MoveFirst
    If rsFCE.EOF Then
        txtCadastro(1).Text = Format(txtCadastro(1), "000000") & ""
        mobjMsg.Abrir "FCE não cadastrada", Ok, critico, "Atenção"
        txtCadastro(1).SetFocus
    Else
        txtCadastro(1).Text = Format(rsFCE.Fields(0), "000000") & ""
        txtCadastro(2).SetFocus
    End If
    rsFCE.Close
    Set rsFCE = Nothing
End Sub

Private Sub ChamaGridFCE()
On Error Resume Next
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
            txtCadastro(1).Text = Format(rsLocal.Fields(0), "000000")
            txtCadastro(2).SetFocus
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Function GeraCodigo()
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
    txtCadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharProjeto
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
