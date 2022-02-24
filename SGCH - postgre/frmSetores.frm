VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de setores"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmSetores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   615
      Left            =   7320
      TabIndex        =   29
      Top             =   7920
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do setor"
      Height          =   3495
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtCadSetor 
         Height          =   1695
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Breve descrição do setor"
         ToolTipText     =   "Breve descrição do setor"
         Top             =   1680
         Width           =   8055
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   4
         Tag             =   "Pesquisar departamento"
         ToolTipText     =   "Pesquisar departamento"
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   3
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
         MICON           =   "frmSetores.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCadSetor 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   3
         Tag             =   "Nome do departamento"
         ToolTipText     =   "Nome do departamento"
         Top             =   1080
         Width           =   6375
      End
      Begin VB.TextBox txtCadSetor 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Tag             =   "Código do departamento"
         ToolTipText     =   "Código do departamento"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCadSetor 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do setor"
         ToolTipText     =   "Código do setor"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadSetor 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do setor"
         ToolTipText     =   "Nome do setor"
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Label9 
         Caption         =   "Nome do departamento:"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Código dep.:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código set.:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nome do setor:"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Histórico de responsáveis pelo setor "
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   8295
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   255
         Index           =   5
         Left            =   7800
         TabIndex        =   8
         Tag             =   "Pesquisar"
         ToolTipText     =   "Pesquisar"
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   3
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
         MICON           =   "frmSetores.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCadSetor 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Tag             =   "Código do responsável pelo setor"
         ToolTipText     =   "Código do responsável pelo setor"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCadSetor 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   7
         Tag             =   "Nome do responsável pelo setor"
         ToolTipText     =   "Nome do responsável pelo setor"
         Top             =   480
         Width           =   6135
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Tag             =   "Excluir responsável"
         ToolTipText     =   "Excluir responsável"
         Top             =   1560
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
         MICON           =   "frmSetores.frx":0D02
         PICN            =   "frmSetores.frx":0D1E
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
         Index           =   2
         Left            =   1320
         TabIndex        =   13
         Tag             =   "Editar responsável"
         ToolTipText     =   "Editar responsável"
         Top             =   1560
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
         MICON           =   "frmSetores.frx":19F8
         PICN            =   "frmSetores.frx":1A14
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
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Tag             =   "Novo responsável"
         ToolTipText     =   "Novo responsável"
         Top             =   1560
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
         MICON           =   "frmSetores.frx":26EE
         PICN            =   "frmSetores.frx":270A
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
         TabIndex        =   11
         Tag             =   "Incluir responsável"
         ToolTipText     =   "Incluir responsável"
         Top             =   1560
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
         MICON           =   "frmSetores.frx":33E4
         PICN            =   "frmSetores.frx":3400
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Tag             =   "Data fim das atividades de responsabilidade pelo setor"
         ToolTipText     =   "Data fim das atividades de responsabilidade pelo setor"
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   156434433
         CurrentDate     =   40549
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Tag             =   "Data início das atividades de responsabilidade pelo setor"
         ToolTipText     =   "Data início das atividades de responsabilidade pelo setor"
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   156434433
         CurrentDate     =   40498
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2990
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin VB.Label Label3 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Data início:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Data fim:"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   7
      Left            =   720
      TabIndex        =   15
      Top             =   7920
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
      MICON           =   "frmSetores.frx":40DA
      PICN            =   "frmSetores.frx":40F6
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
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   7920
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
      MICON           =   "frmSetores.frx":4DD0
      PICN            =   "frmSetores.frx":4DEC
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
Attribute VB_Name = "frmSetores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsSetores As New ADODB.Recordset
Private SqlSetores As String
Private Status As String
Private rsDepartamento As New ADODB.Recordset
Private sqlDepartamento As String
Private rsColaborador As New ADODB.Recordset
Private SqlColaborador As String
Private rsHistorico As New ADODB.Recordset
Private SqlHistorico As String
Private rsLocal As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirResponsavel
        LimpaControlesResp
    Case 1
        LimpaControlesResp
    Case 2
        AlteraResponsavel
    Case 3
        ExcluirItemLV ListView1
    Case 4
        ChamaGridDepartamento
        CarregaDepartamento
    Case 5
        ChamaGridColaborador
        CarregaColaborador
    Case 6
        If MsgBox("Deseja salvar os dados do Setor?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravarDados
            gravaLog "Setor:" & txtCadSetor(0) & "-" & txtCadSetor(1), "Departamento: " & txtCadSetor(2) & "-" & txtCadSetor(3), ""
            'AtivaLD
            Pesquisa = "0"
            Unload Me
            Set frmSetores = Nothing
        End If
    Case 7
        If MsgBox("Deseja sair da tela de cadastro de Setores?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Pesquisa = "0"
            Unload Me
            Set frmSetores = Nothing
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
    listview_cabecalho
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        'DesbloqueiaControles
    End If
    configControles
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Registro", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Responsável", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "Data Início", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Data Fim", ListView1.Width / 6
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarSetor As New ADODB.Recordset
    Dim SqlSalvarSetor As String
    Dim rsDeletarHistResp As New ADODB.Recordset
    Dim SqlDeletarHistResp As String
    Dim rsSalvarHistResp As New ADODB.Recordset
    Dim SqlSalvarHistResp As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarSetor = "select * from tbSetores where codcoligada = '" & vCodcoligada & "' and codSetor = '" & Val(txtCadSetor(0)) & "'"
    rsSalvarSetor.Open SqlSalvarSetor, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarSetor.EOF Then rsSalvarSetor.AddNew
    rsSalvarSetor.Fields(0) = Val(txtCadSetor(0))
    rsSalvarSetor.Fields(1) = txtCadSetor(1)
    rsSalvarSetor.Fields(2) = txtCadSetor(2)
    rsSalvarSetor.Fields(3) = txtCadSetor(4)
    
    If Check1.Value = 0 Then
        rsSalvarSetor.Fields(4) = "N"
    Else
        rsSalvarSetor.Fields(4) = "S"
    End If
    rsSalvarSetor.Fields(5) = vCodcoligada 'Codigo da coligada
    
    rsSalvarSetor.Update
    
    
    SqlDeletarHistResp = "Delete from tbSetoresHistResp where codcoligada = '" & vCodcoligada & "' and codSetor = '" & Val(txtCadSetor(0)) & "' and coddepartamento = '" & Val(txtCadSetor(2)) & "'"
    rsDeletarHistResp.Open SqlDeletarHistResp, cnBanco
      
    SqlSalvarHistResp = "select * from tbSetoresHistResp where codcoligada = '" & vCodcoligada & "'"
    rsSalvarHistResp.Open SqlSalvarHistResp, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvarHistResp.AddNew
        rsSalvarHistResp.Fields(0) = Val(txtCadSetor(0))
        rsSalvarHistResp.Fields(1) = Val(txtCadSetor(2))
        rsSalvarHistResp.Fields(2) = ListView1.ListItems.Item(X)
        rsSalvarHistResp.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2)
        If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then rsSalvarHistResp.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(3)
        rsSalvarHistResp.Fields(4) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvarHistResp.EOF Then rsSalvarHistResp.Update
    
    
    cnBanco.CommitTrans
    rsSalvarHistResp.Close
    Set rsSalvarHistResp = Nothing
    rsSalvarSetor.Close
    Set rsSalvarSetor = Nothing
    MsgBox "Os dados do Setor foram salvos com sucesso", vbInformation, "SGCH"
    AtualizaListview
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker2.Value = ""
    'DesbloqueiaControles
    For X = 0 To txtCadSetor.Count - 1
        txtCadSetor(X) = ""
    Next
    txtCadSetor(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub LimpaControlesResp()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker2.Value = ""
    txtCadSetor(5) = ""
    txtCadSetor(6) = ""
    txtCadSetor(5).Enabled = True
    txtCadSetor(6).Enabled = True
    cmdCadastro(5).Enabled = True
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadSetor(0).Text = Format(rsSetores.Fields(0), "000000")
    txtCadSetor(1).Text = rsSetores.Fields(1)
    txtCadSetor(2).Text = Format(rsSetores.Fields(2), "000000")
    txtCadSetor(3).Text = rsSetores.Fields(6)
    
    If Not IsNull(rsSetores.Fields(3)) Then txtCadSetor(4).Text = rsSetores.Fields(3)
    If rsSetores.Fields(4) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Sub Compoe_Listview()
    ' Declaração de variaveis
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsHistorico.EOF
        Set ItemLst = ListView1.ListItems.Add(, , rsHistorico.Fields(2))
        ItemLst.SubItems(1) = "" & rsHistorico.Fields(5)
        ItemLst.SubItems(2) = "" & rsHistorico.Fields(3)
        If Not IsNull(rsHistorico.Fields(4)) Then ItemLst.SubItems(3) = rsHistorico.Fields(4) Else ItemLst.SubItems(3) = "-"
        rsHistorico.MoveNext
        X = X + 1
    Wend
    'Legenda = ""
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadSetor(0).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadSetor(0).Tag, vbInformation, "Atenção"
        Me.txtCadSetor(0).SetFocus
        Exit Function
    End If
    If txtCadSetor(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadSetor(1).Tag, vbInformation, "Atenção"
        Me.txtCadSetor(1).SetFocus
        Exit Function
    End If
    If txtCadSetor(2).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadSetor(2).Tag, vbInformation, "Atenção"
        Me.txtCadSetor(2).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadSetor.Count - 1
        txtCadSetor(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadSetor.Count - 1
        txtCadSetor(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirSetor
    SqlGera = "Select top 1 * from tbSetores where codcoligada = '" & vCodcoligada & "' order by codSetor Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsSetores.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadSetor(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharSetores
End Function

Private Sub AbrirSetor()
    SqlSetores = "Select * from tbSetores where codcoligada = '" & vCodcoligada & "' Order by codSetor"
    rsSetores.Open SqlSetores, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharSetores()
    rsSetores.Close
    Set rsSetores = Nothing
End Sub

Private Sub AbrirHistorico()
    SqlHistorico = "Select tbSetoresHistResp.*,tbcolaboradores.nomecolaborador from tbSetoresHistResp inner join tbcolaboradores on tbcolaboradores.codcoligada = '" & vCodcoligada & "' and tbcolaboradores.codcolaborador = tbSetoresHistResp.codcolaborador where codSetor = '" & Val(txtCadSetor(0)) & "' and coddepartamento = '" & Val(txtCadSetor(2)) & "'"
    rsHistorico.Open SqlHistorico, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharHistorico()
    rsHistorico.Close
    Set rsHistorico = Nothing
End Sub

Private Sub ResultPesq()
    SqlSetores = "Select tbsetores.*,tbdepartamentos.nomedepartamento from tbSetores,tbdepartamentos Where tbSetores.codcoligada = '" & vCodcoligada & "' and tbSetores.codSetor= '" & Val(varGlobal) & "' and tbdepartamentos.coddepartamento = tbsetores.coddepartamento order by tbsetores.codSetor"
    rsSetores.Open SqlSetores, cnBanco, adOpenKeyset, adLockReadOnly
    If rsSetores.RecordCount > 0 Then
        CompoeControles
        AbrirHistorico
        Compoe_Listview
        FecharHistorico
    Else
        MsgBox "Setor não encontrado"
    End If
    rsSetores.Close
    Set rsSetores = Nothing
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
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadSetor(0), "000000"))
        ItemLst.SubItems(1) = txtCadSetor(1).Text
        ItemLst.SubItems(2) = Format(txtCadSetor(2).Text, "000000") & " - " & txtCadSetor(3).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadSetor(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = Format(txtCadSetor(2).Text, "000000") & " - " & txtCadSetor(3).Text
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
    MsgBox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraResponsavel
    End If
End Sub

Private Sub txtCadSetor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Error
    Select Case Index
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaDepartamento
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub CarregaDepartamento()
    Dim X As Integer
    sqlDepartamento = "Select * from tbdepartamentos where codcoligada = '" & vCodcoligada & "' and ativo <> 'N' order by coddepartamento"
    rsDepartamento.Open sqlDepartamento, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsDepartamento.EOF Then rsDepartamento.MoveFirst
    rsDepartamento.Find "coddepartamento=" & "'" & Val(Me.txtCadSetor(2)) & "'"
    If rsDepartamento.EOF Then
        txtCadSetor(2).Text = Format(txtCadSetor(2), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            MsgBox "Departamento não cadastrado", vbInformation, "SGCH"
            txtCadSetor(3) = ""
            cbocad = ""
        End If
    Else
        txtCadSetor(2).Text = Format(rsDepartamento.Fields(0), "000000") & ""
        txtCadSetor(3).Text = rsDepartamento.Fields(1)
    End If
    rsDepartamento.Close
    Set rsDepartamento = Nothing
End Sub

Private Sub ChamaGridDepartamento()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbdepartamentos where codcoligada = '" & vCodcoligada & "' and ativo <> 'N' order by nomedepartamento"
    procnom = "nomedepartamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de departamento"
    Pesquisa = frmSetores.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomedepartamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadSetor(2).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub CarregaColaborador()
    Dim X As Integer
    SqlColaborador = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' and ativo = 'S' order by codcolaborador"
    rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsColaborador.EOF Then rsColaborador.MoveFirst
    rsColaborador.Find "codcolaborador=" & "'" & Me.txtCadSetor(5) & "'"
    If rsColaborador.EOF Then
        txtCadSetor(5).Text = txtCadSetor(5)
        If Val(Pesquisa) <> 0 Then
            MsgBox "Colaborador não cadastrado", vbInformation, "SGCH"
            txtCadSetor(6) = ""
        End If
    Else
        txtCadSetor(5).Text = rsColaborador.Fields(1)
        txtCadSetor(6).Text = rsColaborador.Fields(3)
    End If
    rsColaborador.Close
    Set rsColaborador = Nothing
End Sub

Private Sub ChamaGridColaborador()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' and ativo = 'S' order by nomecolaborador"
    procnom = "nomecolaborador"
    campo = 3
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Colaborador"
    Pesquisa = frmSetores.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecolaborador=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadSetor(5).Text = rsLocal.Fields(1)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

' (INICIO) >>>>>>>> CONTROLES DOS BOTOES DE HISTÓRICO DE RESPONSÁVEIS <<<<<<<<<<
Private Sub IncluirResponsavel()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaResponsavel = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtCadSetor(5) Then
                Me.txtCadSetor(5) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadSetor(6)
                ListView1.SelectedItem.ListSubItems.Item(2) = DTPicker1
                If DTPicker2.Value <> "" Then ListView1.SelectedItem.ListSubItems.Item(3) = DTPicker2 Else ListView1.SelectedItem.ListSubItems.Item(3) = "-"
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtCadSetor(5))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtCadSetor(5))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadSetor(6)
    ItemLst.SubItems(2) = DTPicker1
    If DTPicker2.Value <> "" Then ItemLst.SubItems(3) = DTPicker2 Else ItemLst.SubItems(3) = "-"
    txtCadSetor(5).SetFocus
End Sub

Private Function ValidaResponsavel()
    ValidaResponsavel = False
    If txtCadSetor(5).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadSetor(5).Tag, vbInformation, "Atenção"
        Me.txtCadSetor(5).SetFocus
        Exit Function
    End If
    ValidaResponsavel = True
End Function

Private Sub AlteraResponsavel()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadSetor(5).Text = ListView1.ListItems.Item(X)
    Me.txtCadSetor(6).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    DTPicker1 = ListView1.SelectedItem.ListSubItems.Item(2)
    If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then DTPicker2 = ListView1.SelectedItem.ListSubItems.Item(3)
    txtCadSetor(5).Enabled = False
    txtCadSetor(6).Enabled = False
    cmdCadastro(5).Enabled = False
End Sub
' (FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    
    End If
    If vEdi = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then
        cmdCadastro(6).UseGreyscale = True
        cmdCadastro(6).DragMode = 1
        cmdCadastro(6).SpecialEffect = cbEngraved
    End If
    If vExc = "N" Then
        cmdCadastro(3).UseGreyscale = True
        cmdCadastro(3).DragMode = 1
        cmdCadastro(3).SpecialEffect = cbEngraved
    End If
End Sub

