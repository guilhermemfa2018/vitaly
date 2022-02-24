VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpesqger 
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   7080
   Icon            =   "frmpesqger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Filtro da pesquisa "
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtPesquisa 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdPesquisa 
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPesquisa 
      Caption         =   "&Ok"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPesquisa 
      Caption         =   "&Pesquisar"
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3855
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   -2147483624
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmpesqger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsTabela As New ADODB.Recordset
Dim sql As String

Private Sub aicAlphaImage1_Click(ByVal Button As Integer)
'    If Pesquisa = "Lista de Materiais" Then frmTipoMat.Show 1
'    If Pesquisa = "Funcionários" Then frmFuncionarios.Show 1
'    If Pesquisa = "Emissão de relatórios" Then frmTransportes.Show 1
End Sub

Private Sub cmdPesquisa_Click(Index As Integer)
    If Index = 0 Then
        If Grid.RowSel > 0 Then
            If procnom = "codrequisicao" Then
                Pesquisa = Grid.TextMatrix(Grid.RowSel, 1) & Grid.TextMatrix(Grid.RowSel, 2)
            ElseIf procnom = "codmatriz" Or procnom = "codavaliacao" Then
                Pesquisa = Grid.TextMatrix(Grid.RowSel, 1)
            ElseIf procnom = "nomecargo" Then
                Pesquisa = Grid.TextMatrix(Grid.RowSel, 1) & Grid.TextMatrix(Grid.RowSel, 2)
            Else
                Pesquisa = Grid.TextMatrix(Grid.RowSel, 2)
            End If
        Else
            Pesquisa = 0
        End If
        Unload Me
        Set frmpesqger = Nothing
    ElseIf Index = 1 Then
        Pesquisa = 0
        Unload Me
        Set frmpesqger = Nothing
    ElseIf Index = 2 Then
        rsTabela.Close
        Set rsTabela = Nothing
        Form_Load
        txtPesquisa(0).SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim sql As String
    sql = Sqlp
    rsTabela.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTabela.RecordCount = 0 Then Exit Sub
    If txtPesquisa(0) <> "" Then
        rsTabela.MoveFirst
        If procnom <> "codmatriz" Then rsTabela.Find procnom & " like '" & txtPesquisa(0).Text & "*'" ' Realiza pesquisa com as iniciais do textbox
        If procnom = "codmatriz" Then rsTabela.Find procnom1 & " like '" & txtPesquisa(0).Text & "*'" ' Realiza pesquisa com as iniciais do textbox
        If rsTabela.EOF Then Exit Sub
    Else
        If rsTabela.RecordCount <> 0 Then
            rsTabela.MoveFirst
        End If
    End If
    
    'USADO NO FORM DE ADMISSÃO DE CANDIDATOS
    If Pesquisa = "Admissao" Then
        'Pesquisa = 0
        Grid.Rows = Grid.FixedRows
        Grid.Cols = 4
        Grid.ColWidth(0) = 200
        
        Grid.ColWidth(1) = 800
        Grid.TextMatrix(0, 1) = "Requisição"
        
        Grid.ColWidth(2) = 800
        Grid.TextMatrix(0, 2) = "Matriz"
        
        Grid.ColWidth(3) = 2000
        Grid.TextMatrix(0, 3) = "Cargo nome"
        
        If rsTabela.RecordCount > 0 Then
            Grid.Rows = Grid.Rows + rsTabela.RecordCount
            X = 0
            Do While Not rsTabela.EOF
                Grid.TextMatrix(X + 1, 1) = Format(rsTabela.Fields(campo), "000000") 'rsTab.Fields(0)
                Grid.TextMatrix(X + 1, 2) = Format(rsTabela.Fields(Campo1), "000000") 'rsTab.Fields(0)
                Grid.TextMatrix(X + 1, 3) = rsTabela.Fields(campo2)
                rsTabela.MoveNext
                X = X + 1
            Loop
        End If
    'USADO EM TODA PESQUISA DE MATRIZ
    ElseIf Pesquisa = "Histórico" Then
        'Pesquisa = 0
        Grid.Rows = Grid.FixedRows
        Grid.Cols = 6
        Grid.ColWidth(0) = 200
        
        Grid.ColWidth(1) = 800
        Grid.TextMatrix(0, 1) = "Matriz"
        
        Grid.ColWidth(2) = 3500
        Grid.TextMatrix(0, 2) = "Cargo"
        
        Grid.ColWidth(3) = 800
        Grid.TextMatrix(0, 3) = "Nível"
        
        Grid.ColWidth(4) = 2000
        Grid.TextMatrix(0, 4) = "Departamento"
        
        Grid.ColWidth(5) = 2000
        Grid.TextMatrix(0, 5) = "Setor"
        Me.Grid.ColAlignment(1) = flexAlignLeftCenter
    
        If rsTabela.RecordCount > 0 Then
            Grid.Rows = Grid.Rows + rsTabela.RecordCount
            X = 0
            Do While Not rsTabela.EOF
                Grid.TextMatrix(X + 1, 1) = Format(rsTabela.Fields(campo), "000000") 'rsTab.Fields(0)
                Grid.TextMatrix(X + 1, 2) = rsTabela.Fields(Campo1)
                Grid.TextMatrix(X + 1, 3) = rsTabela.Fields(campo2)
                Grid.TextMatrix(X + 1, 4) = rsTabela.Fields(campo3)
                Grid.TextMatrix(X + 1, 5) = rsTabela.Fields(Campo4)
                rsTabela.MoveNext
                X = X + 1
            Loop
        End If
    'USADO EM TODA CONSULTA GENERICA
    Else
        Grid.Rows = Grid.FixedRows
        Grid.Cols = 3
        Grid.ColWidth(0) = 200
        Grid.ColWidth(1) = 1500
        Grid.TextMatrix(0, 1) = "Código"
        Grid.ColWidth(2) = 4800
        Grid.TextMatrix(0, 2) = "Descrição"
        Me.Grid.ColAlignment(2) = flexAlignLeftCenter
        If rsTabela.RecordCount > 0 Then
            Grid.Rows = Grid.Rows + rsTabela.RecordCount
            X = 0
            Do While Not rsTabela.EOF
                If procnom <> "codorc" And procnom <> "nomecolaborador" Then
                    Grid.TextMatrix(X + 1, 1) = Format(rsTabela.Fields(Campo1), "000000") 'rsTab.Fields(0)
                Else
                    Grid.TextMatrix(X + 1, 1) = rsTabela.Fields(Campo1) 'rsTab.Fields(0)
                End If
                Grid.TextMatrix(X + 1, 2) = rsTabela.Fields(campo)
                rsTabela.MoveNext
                X = X + 1
            Loop
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsTabela.Close
    Set rsTabela = Nothing
    Set frmpesqger = Nothing
End Sub

Private Sub Grid_DblClick()
    If Grid.RowSel > 0 Then
        If procnom = "codrequisicao" Then
            Pesquisa = Grid.TextMatrix(Grid.RowSel, 1) & Grid.TextMatrix(Grid.RowSel, 2)
        ElseIf procnom = "codmatriz" Or procnom = "codavaliacao" Then
            Pesquisa = Grid.TextMatrix(Grid.RowSel, 1)
        ElseIf procnom = "nomecargo" Then
            Pesquisa = Grid.TextMatrix(Grid.RowSel, 1) & Grid.TextMatrix(Grid.RowSel, 2)
        Else
            Pesquisa = Grid.TextMatrix(Grid.RowSel, 2)
        End If
    Else
        Pesquisa = 0
    End If
    Unload Me
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If Grid.RowSel > 0 Then
        Pesquisa = Grid.TextMatrix(Grid.RowSel, 2)
    Else
        Pesquisa = 0
    End If
    'MsgBox Pesquisa
    Unload Me
    Set frmpesqger = Nothing
End Sub

Private Sub txtPesquisa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        rsTabela.Close
        Set rsTabela = Nothing
        Form_Load
        'txtPesquisa(0).SetFocus
        Grid.SetFocus
    End If
End Sub
