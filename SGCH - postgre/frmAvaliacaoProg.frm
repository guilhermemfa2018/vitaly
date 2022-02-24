VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAvaliacaoProg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Itens de avaliação de eficácia do treinamento "
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "frmAvaliacaoProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   10
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   6960
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
      MICON           =   "frmAvaliacaoProg.frx":0CCA
      PICN            =   "frmAvaliacaoProg.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Tag             =   "Carregar MODELO"
      ToolTipText     =   "Carregar MODELO"
      Top             =   6960
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
      MICON           =   "frmAvaliacaoProg.frx":19C0
      PICN            =   "frmAvaliacaoProg.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modelos de avaliação"
      Height          =   6735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      Begin SGCH.chameleonButton cmdModelo 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Excluir modelo"
         ToolTipText     =   "Excluir modelo"
         Top             =   1440
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
         MICON           =   "frmAvaliacaoProg.frx":26B6
         PICN            =   "frmAvaliacaoProg.frx":26D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Tag             =   "Nome do modelo de avaliação"
         ToolTipText     =   "Nome do modelo de avaliação"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtModelo 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7858
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SGCH.chameleonButton cmdModelo 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Tag             =   "Editar modelo"
         ToolTipText     =   "Editar modelo"
         Top             =   1440
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
         MICON           =   "frmAvaliacaoProg.frx":33AC
         PICN            =   "frmAvaliacaoProg.frx":33C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdModelo 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Tag             =   "Novo modelo"
         ToolTipText     =   "Novo modelo"
         Top             =   1440
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
         MICON           =   "frmAvaliacaoProg.frx":40A2
         PICN            =   "frmAvaliacaoProg.frx":40BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdModelo 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "Incluir modelo"
         ToolTipText     =   "Incluir modelo"
         Top             =   1440
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
         MICON           =   "frmAvaliacaoProg.frx":4D98
         PICN            =   "frmAvaliacaoProg.frx":4DB4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Modelo ativo:"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Marque os itens que serão avaliados nesse treinamento"
      Height          =   6735
      Left            =   4080
      TabIndex        =   11
      Top             =   120
      Width           =   4575
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   11033
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   6960
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
      MICON           =   "frmAvaliacaoProg.frx":5A8E
      PICN            =   "frmAvaliacaoProg.frx":5AAA
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
Attribute VB_Name = "frmAvaliacaoProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdModelo_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirModelo
        LimpaControlesModelo
        desmarcaTodos
        desmarcaModelos
    Case 1
        LimpaControlesModelo
        CompoeAE
        desmarcaTodos
    Case 2
        editaControlesModelo
        desmarcaTodos
    End Select
End Sub

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        GravarDados
        vCodModeloAval = Val(txtModelo(2))
    Case 1
        Unload Me
        Set frmAvaliacaoProg = Nothing
    Case 2
        vCodModeloAval = Val(txtModelo(2))
        chamaForm.Label25 = Format(vCodModeloAval, "000")
    End Select
End Sub

Private Sub Form_Load()
    listview_cabecalho
    txtModelo(2) = vCodModeloAval
    compoeLV2
    marcaLV2
    LimpaControlesModelo
    CompoeAE
    'CompoePontosAE
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Avaliação", ListView1.Width / 1.5
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 1.5
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub compoeLV2()
    Dim rsModelo As New ADODB.Recordset
    Dim sqlModelo As String
    sqlModelo = "select * from tbModeloProg where codcoligada ='" & vCodcoligada & "'"
    rsModelo.Open sqlModelo, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView2.ListItems.Clear
    While Not rsModelo.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsModelo.Fields(0), "000")) 'codigo do modelo de avaliação
        ItemLst.SubItems(1) = "" & rsModelo.Fields(1) 'nome do modelo de avaliação
        rsModelo.MoveNext
        X = X + 1
    Wend
    rsModelo.Close
    Set rsModelo = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub marcaLV2()
    If txtModelo(2) = "" Then Exit Sub
    Dim X As Integer, Y As Integer, J As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If Val(ListView2.ListItems.Item(X)) = Val(txtModelo(2)) Then
            ListView2.ListItems.Item(X).Checked = True
            Exit Sub
        End If
    Next
End Sub

Private Sub CompoeAE()
    Dim rsAE As New ADODB.Recordset
    Dim sqlAE As String
    'Tabela que armazena os itens que serão avaliados na Avaliação de Eficácia do Treinamento
    sqlAE = "select * from tbAvaliacao where codcoligada = '" & vCodcoligada & "' and tipo = 'AE' and ativo = 'S'"
    rsAE.Open sqlAE, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    ListView1.ListItems.Clear
    While Not rsAE.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsAE.Fields(0), "000")) 'codigo da avaliação AE
        ItemLst.SubItems(1) = "" & rsAE.Fields(1) 'nome da avaliação AE
        rsAE.MoveNext
        X = X + 1
    Wend
    rsAE.Close
    Set rsAE = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE PONTUAÇÃO DA AVALIAÇÃO DE EFICÁCIA <<<<<<<<<<
Private Sub CompoePontosAE()
    Dim rsHab As New ADODB.Recordset
    Dim sqlHab As String
    sqlHab = "Select * from tbAvaliacaoProg where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & Val(txtModelo(2)) & "'"
    rsHab.Open sqlHab, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If rsHab.RecordCount = 0 Then Exit Sub
    Y = ListView1.ListItems.Count
    While Not rsHab.EOF
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True
            If Val(ListView1.ListItems.Item(X)) = rsHab.Fields(0) Then
                ListView1.ListItems.Item(X).Checked = True
            End If
        Next
        rsHab.MoveNext
    Wend
    rsHab.Close
    Set rsHab = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    Dim Y As Integer, X As Integer, K As Integer
    cnBanco.BeginTrans
    '>>>>>> GRAVAR AVALIACAO <<<<<<<<<
    sqlDeletar = "Delete from tbModeloProg where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbModeloProg"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    If ListView2.ListItems.Count > 0 Then
        For X = 1 To ListView2.ListItems.Count
            ListView2.ListItems.Item(X).Selected = True
            rsSalvar.AddNew
            rsSalvar.Fields(0) = Val(ListView2.ListItems.Item(X))
            rsSalvar.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
            rsSalvar.Fields(2) = vCodcoligada 'Codigo da Coligada
        Next
    End If
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    sqlDeletar = "Delete from tbAvaliacaoProg where codmodelo = '" & Val(txtModelo(2)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbAvaliacaoProg where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & Val(txtModelo(2)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    If ListView1.ListItems.Count > 0 Then
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.ListItems.Item(X).Checked = True Then
                rsSalvar.Find "codavaliacao=" & "'" & Val(ListView1.ListItems.Item(X)) & "'"
                If rsSalvar.EOF Then rsSalvar.AddNew
                rsSalvar.Fields(0) = Val(ListView1.ListItems.Item(X))
                rsSalvar.Fields(1) = Val(txtModelo(2))
                rsSalvar.Fields(2) = vCodcoligada 'Codigo da coligada
            End If
        Next
    End If
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    cnBanco.CommitTrans
    MsgBox "Os dados foram salvos com sucesso", vbInformation, "SGCH"
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
TrataErro1:
    Resume Next
End Sub

'-------- ROTINAS DE MODELOS DE AVALIAÇÃO DE EFICÁCIA -------------
Private Sub IncluirModelo()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaModelo = False Then Exit Sub
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView2.ListItems.Item(X) = Me.txtModelo(0) Then
                Me.txtModelo(0) = ListView2.ListItems.Item(X)
                ListView2.SelectedItem.ListSubItems.Item(1) = txtModelo(1)
                Y = ListView2.ListItems.Count
                Me.ListView2.SortOrder = lvwAscending
                Exit Sub
            End If
        Next
        Set ItemLst = ListView2.ListItems.Add(, , txtModelo(0))
        Y = ListView2.ListItems.Count
    Else
        Set ItemLst = ListView2.ListItems.Add(, , txtModelo(0))
        Y = ListView2.ListItems.Count
        Me.ListView2.SortOrder = lvwAscending
    End If
    ItemLst.SubItems(1) = txtModelo(1)
    txtModelo(1).SetFocus
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Function ValidaModelo()
    ValidaModelo = False
    If txtModelo(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtModelo(1).Tag, vbInformation, "Atenção"
        Me.txtModelo(1).SetFocus
        Exit Function
    End If
    ValidaModelo = True
End Function

Private Sub LimpaControlesModelo()
    Dim X As Integer
    txtModelo(1) = ""
    If ListView2.ListItems.Count > 0 Then
        txtModelo(0).Text = Format(GeraCodigo1, "000")
    Else
        txtModelo(0).Text = Format(Val(txtModelo(0)) + 1, "000")
    End If
End Sub

Private Function GeraCodigo1()
    Dim X As Integer
    X = 1
    Me.ListView2.SortOrder = lvwDescending
    ListView2.ListItems.Item(X).Selected = True
    GeraCodigo1 = Val(ListView2.ListItems.Item(X)) + 1
    Me.ListView2.SortOrder = lvwAscending
    Exit Function
End Function

Private Sub ListView2_Click()
   'pegaClick
   desmarcaModelos
   CompoeAE
   CompoePontosAE
End Sub

Private Sub desmarcaModelos()
    Dim X As Integer, Y As Integer, J As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    J = ListView2.SelectedItem.Index
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Item(X).Checked = False
        End If
    Next
    ListView2.ListItems.Item(J).Checked = True
    If ListView2.ListItems.Item(J).Checked = True Then ListView1.Enabled = True
    txtModelo(2) = ListView2.ListItems.Item(J)
End Sub

Private Sub desmarcaTodos()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        ListView2.ListItems.Item(X).Checked = False
    Next
    ListView1.Enabled = False
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ListView2.ListItems(Item.Index).Selected = True
End Sub

Private Sub ListView2_DblClick()
    editaControlesModelo
End Sub

Private Sub editaControlesModelo()
    If ListView2.ListItems.Count = 0 Then Exit Sub
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtModelo(0).Text = ListView2.ListItems.Item(X)
    Me.txtModelo(1).Text = ListView2.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub excluirControlesModelo()
    Dim rsVSePodeExcluir As New ADODB.Recordset
    Dim SqlVSePodeExcluir As String
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    Dim llng_Contador As Long
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    SqlVSePodeExcluir = "Select codmodelo from tbprogramacao where codcoligada = '" & vCodcoligada & "' and codmodelo = '" & ListView2.ListItems.Item(X) & "'"
    rsVSePodeExcluir.Open SqlVSePodeExcluir, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVSePodeExcluir.RecordCount > 0 Then
        MsgBox "Existem programações para esse modelo. Não pode ser excluido", vbCritical, "SGCH"
        Exit Sub
    End If
    ListView2.ListItems.Remove (X)
End Sub

Private Sub txtModelo_Change(Index As Integer)
    Select Case Index
    Case 1
        CompoeAE
        desmarcaTodos
    End Select
End Sub

Private Sub txtModelo_Click(Index As Integer)
    Select Case Index
    Case 1
        CompoeAE
        desmarcaTodos
    End Select
End Sub

Private Sub txtModelo_GotFocus(Index As Integer)
    Select Case Index
    Case 1
        CompoeAE
        desmarcaTodos
    End Select
End Sub
