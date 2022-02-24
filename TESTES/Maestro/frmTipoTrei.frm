VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTipoTrei 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Treinamento"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmTipoTrei.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
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
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6975
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5318
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do tipo de treinamento"
         ToolTipText     =   "Nome do tipo de treinamento"
         Top             =   480
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmTipoTrei.frx":0CCA
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do tipo de treinamento"
         ToolTipText     =   "Código do tipo de treinamento"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTipoTrei.frx":0D32
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Excluir tipo de treinamento"
         ToolTipText     =   "Excluir tipo de treinamento"
         Top             =   840
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
         MICON           =   "frmTipoTrei.frx":0D9E
         PICN            =   "frmTipoTrei.frx":0DBA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Tag             =   "Editar tipo de treinamento"
         ToolTipText     =   "Editar tipo de treinamento"
         Top             =   840
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
         MICON           =   "frmTipoTrei.frx":1A94
         PICN            =   "frmTipoTrei.frx":1AB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Tag             =   "Novo tipo de treinamento"
         ToolTipText     =   "Novo tipo de treinamento"
         Top             =   840
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
         MICON           =   "frmTipoTrei.frx":278A
         PICN            =   "frmTipoTrei.frx":27A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "Incluir tipo de treinamento"
         ToolTipText     =   "Incluir tipo de treinamento"
         Top             =   840
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
         MICON           =   "frmTipoTrei.frx":3480
         PICN            =   "frmTipoTrei.frx":349C
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
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   720
      TabIndex        =   8
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   4920
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
      MICON           =   "frmTipoTrei.frx":4176
      PICN            =   "frmTipoTrei.frx":4192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   4920
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
      MICON           =   "frmTipoTrei.frx":4E6C
      PICN            =   "frmTipoTrei.frx":4E88
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
Attribute VB_Name = "frmTipoTrei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTipoTrei As New ADODB.Recordset
Private sqlTipoTrei As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirTipo
        LimpaControles
    Case 1
        LimpaControles
    Case 2
        AlteraTipo
    Case 3
        ExcluirItemLV ListView1
    Case 4
        GravarDados
        Unload Me
    Case 5
        Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    listview_cabecalho
    Compoe_Listview
    LimpaControles
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
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 3
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X) = ""
    Next
    txtCadastro(0) = Format(GeraCodigo(ListView1), "000000")
End Sub

Private Sub AbrirTipoTrei()
    sqlTipoTrei = "Select * from tbTipoTrei where codcoligada = '" & vCodcoligada & "' Order by codigo"
    rsTipoTrei.Open sqlTipoTrei, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharTipoTrei()
    rsTipoTrei.Close
    Set rsTipoTrei = Nothing
End Sub

Private Function GeraCodigo(LV As ListView)
    If LV.ListItems.Count > 0 Then
        Dim X As Integer
        X = 1
        LV.SortOrder = lvwDescending
        LV.ListItems.Item(X).Selected = True
        GeraCodigo = LV.ListItems.Item(X) + 1
        LV.SortOrder = lvwAscending
        Exit Function
    Else
        GeraCodigo = 1
    End If
End Function

Private Sub Compoe_Listview()
    ' Declaração de variaveis
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    AbrirTipoTrei
    While Not rsTipoTrei.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsTipoTrei.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsTipoTrei.Fields(1)
        rsTipoTrei.MoveNext
        X = X + 1
    Wend
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
    FecharTipoTrei
End Sub

' (INICIO) >>>>>>>> CONTROLES DOS BOTOES DE HISTÓRICO DE RESPONSÁVEIS DO DEPARTAMENTO<<<<<<<<<<
Private Sub IncluirTipo()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaTipo = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtCadastro(0) Then
                Me.txtCadastro(0) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadastro(1)
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtCadastro(0))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtCadastro(0))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadastro(1)
    txtCadastro(1).SetFocus
End Sub

Private Function ValidaTipo()
    ValidaTipo = False
    Dim X As Integer
    For X = 0 To txtCadastro.Count - 1
        If txtCadastro(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtCadastro(X).Tag, Ok, critico, "Atenção"
            Me.txtCadastro(X).SetFocus
            Exit Function
        End If
    Next
    ValidaTipo = True
End Function

Private Sub AlteraTipo()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadastro(0).Text = ListView1.ListItems.Item(X)
    Me.txtCadastro(1).Text = ListView1.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraTipo
    End If
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
    
    sqlDeletar = "Delete from tbTipoTrei where codcoligada ='" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
      
    sqlTipoTrei = "select * from tbTipoTrei"
    rsTipoTrei.Open sqlTipoTrei, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsTipoTrei.AddNew
        rsTipoTrei.Fields(0) = ListView1.ListItems.Item(X)
        rsTipoTrei.Fields(1) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsTipoTrei.Fields(2) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsTipoTrei.EOF Then rsTipoTrei.Update
    rsTipoTrei.Close
    Set rsTipoTrei = Nothing

    cnBanco.CommitTrans
    mobjMsg.Abrir "Os dados dos Tipos de treinamentos foram salvos com sucesso", Ok, informacao, "Atenção"
    'AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

