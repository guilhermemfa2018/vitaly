VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProcSelAddDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informa��es complementares do candidato"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "frmProcSelAddDados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      Caption         =   "ATEN��O "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   4815
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmProcSelAddDados.frx":0CCA
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Integra��o Totvs"
      Height          =   855
      Left            =   5760
      TabIndex        =   24
      Tag             =   "Verifique se os dados de integra��o no cadastro do colaborador est�o corretamente preenchidos"
      ToolTipText     =   "Verifique se os dados de integra��o no cadastro do colaborador est�o corretamente preenchidos"
      Top             =   3480
      Width           =   4695
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   9
         Left            =   960
         TabIndex        =   26
         Tag             =   "Fun��o"
         ToolTipText     =   "Fun��o"
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtCons 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Tag             =   "Fun��o"
         ToolTipText     =   "Fun��o"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblCons 
         Caption         =   "Fun��o:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cargo da Requisi��o "
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   5895
      Begin VB.Label Label5 
         Caption         =   "Cargo Requisitado"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Registro n�: "
      Height          =   615
      Left            =   6120
      TabIndex        =   15
      Top             =   1920
      Width           =   2295
      Begin VB.TextBox txtNovoCol 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Tag             =   "Registro do novo colaborador"
         ToolTipText     =   "Registro do novo colaborador"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Observa��o"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   7695
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmProcSelAddDados.frx":0D2B
         Left            =   120
         List            =   "frmProcSelAddDados.frx":0D32
         TabIndex        =   14
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Status"
      Height          =   735
      Left            =   7920
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informa��es do candidato "
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6120
         TabIndex        =   23
         Text            =   "Text3"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Tag             =   "Matriz e cargo do colaborador"
         Text            =   "matriz - cargo"
         ToolTipText     =   "Matriz e cargo do colaborador"
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNovoColaborador 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   600
         Width           =   4335
      End
      Begin VB.Frame Frame10 
         Caption         =   "M�dia geral"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   735
         Left            =   6840
         TabIndex        =   3
         Top             =   120
         Width           =   1335
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   60
            TabIndex        =   4
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Label Label17 
         Caption         =   "ID Colaborador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "CPF n�:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label42 
         Caption         =   "Matriz/Cargo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Foto "
      Height          =   2415
      Index           =   0
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   1635
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   2175
            Left            =   0
            Top             =   -120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   3836
            Image           =   "frmProcSelAddDados.frx":0D51
         End
      End
      Begin MSComDlg.CommonDialog cdlFoto 
         Left            =   1080
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   19
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   4440
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
      MICON           =   "frmProcSelAddDados.frx":0D69
      PICN            =   "frmProcSelAddDados.frx":0D85
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
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Tag             =   "Confirmar"
      ToolTipText     =   "Confirmar"
      Top             =   4440
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
      MICON           =   "frmProcSelAddDados.frx":1A5F
      PICN            =   "frmProcSelAddDados.frx":1A7B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
      Height          =   600
      Left            =   5040
      Top             =   3720
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Image           =   "frmProcSelAddDados.frx":2755
      Props           =   5
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   9015
   End
End
Attribute VB_Name = "frmProcSelAddDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsNovoColaboradores As New ADODB.Recordset
Private SqlNovoColaboradores As String

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        If ValidaCampo = False Then Exit Sub
        If vIntegra = "S" Then
            AddDadosGeral(7) = txtCons(8)
        End If
        AddDadosGeral(8) = Combo1
        AddDadosGeral(9) = txtNovoCol(1)
        If Label9 = "Aprovado" Then
            AddDadosGeral(8) = "-"
        End If
        Sqlp = True
    Case 2
        AddDadosGeral(8) = "-"
        AddDadosGeral(9) = ""
        Sqlp = False
    End Select
    Unload Me
    Set frmProcSelAddDados = Nothing
End Sub

Private Sub Form_Load()
    Sqlp = False
    ResultPesq
    If vIntegra = "S" Then
        Frame11.Visible = True
        Frame12.Visible = True
        aicAlphaImage2.Visible = True
    Else
        Frame11.Visible = False
        Frame12.Visible = False
        aicAlphaImage2.Visible = False
    End If
    'configControles
    If vIntegra = "S" Then ConexaoTotvs
    If vIntegra = "S" Then comporCombosTotvs
    If vIntegra = "S" Then
        comporControlesTotvs
    End If
End Sub

Private Sub ResultPesq()
    txtNovoColaborador(0) = AddDadosGeral(0)
    txtNovoColaborador(1) = AddDadosGeral(1)
    txtNovoColaborador(2) = AddDadosGeral(4)
    txtNovoColaborador(3) = AddDadosGeral(2) & " - " & AddDadosGeral(3)
    txtNovoCol(1) = AddDadosGeral(9)
    Label5 = AddDadosGeral(5)
    Label41 = AddDadosGeral(6)
    If Val(Label41) < MediaGlobal And Val(Label41) >= vAprovadoRest Then
        Label41.ForeColor = &H40C0&
        Label9.ForeColor = &H40C0&
        Label9.Caption = "Aprovado com restri��o"
    ElseIf Val(Label41) < vAprovadoRest Then
        Label41.ForeColor = &HC0&
        Label9.ForeColor = &HC0&
        Label9.Caption = "Reprovado"
    ElseIf Val(Label41) >= MediaGlobal Then
        Label41.ForeColor = &H8000&
        Label9.ForeColor = &H8000&
        Frame4.Enabled = False
        Combo1.Enabled = False
        Label9.Caption = "Aprovado"
    End If
    Combo1 = AddDadosGeral(7)
    
    If AddDadosGeral(4) = "candidato" Then 'Candidato
        SqlNovoColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.id from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'candidato' and b.ativo = 'S' and a.cpf = '" & AddDadosGeral(0) & "' order by a.cpf"
    ElseIf AddDadosGeral(4) = "colaborador" Then 'Colaborador
        SqlNovoColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.codcolaborador,a.id from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'colaborador' and b.ativo = 'S' and a.cpf = '" & AddDadosGeral(0) & "' order by a.cpf"
    End If
    rsNovoColaboradores.Open SqlNovoColaboradores, cnBanco, adOpenKeyset, adLockReadOnly

    If AddDadosGeral(4) = "colaborador" Then
        txtNovoCol(1) = rsNovoColaboradores.Fields(5)
        txtNovoCol(1).Enabled = False
        Label17 = rsNovoColaboradores.Fields(6)
    Else
        Label17 = rsNovoColaboradores.Fields(5)
    End If
    
    If Not rsNovoColaboradores.EOF Then Label53.Caption = rsNovoColaboradores.Fields(4)
    If Not rsNovoColaboradores.EOF Then aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    rsNovoColaboradores.Close
    Set rsNovoColaboradores = Nothing
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If avaliaAdmissao = False Then Exit Function
    
    If vIntegra = "S" Then
        If Combo(9).Text = "" Then
            MsgBox "Os dados de integra��o Totvs devem ser informados"
            Exit Function
        End If
    End If
    If Label9 <> "Aprovado" And Combo1 = "" Then
        MsgBox "Deve ser apresentado uma justificativa para admiss�o do candidato"
        Exit Function
    End If
    If txtNovoCol(1) = "" Or txtNovoCol(1) = "-" Then
        MsgBox "Favor informar o registro do novo colaborador"
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function avaliaAdmissao()
    avaliaAdmissao = False
'-Padrao - para saber se ja tem uma solicita��o cadastrada --------------------------------
    Dim vNumPDO As Integer
    Dim vStatusPDO As String
    Dim vDecisao As String
    Dim rsPDOColab As New ADODB.Recordset
    Dim SqlPDOColab As String
   
    SqlPDOColab = "Select a.cpf,a.codcolaborador,a.nomecolaborador,b.id,b.status,b.tipo,b.decisao,a.datarecisao from tbcolaboradores as a left join tbautorizacao as b on a.autorizacao = b.id where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & txtNovoColaborador(0) & "'"
    rsPDOColab.Open SqlPDOColab, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not IsNull(rsPDOColab.Fields(7)) Then
        MsgBox "Colaborador DEMITIDO, n�o pode ser admitido atrav�s desse m�dulo"
        Exit Function
    End If
    
    If Not IsNull(rsPDOColab.Fields(3)) Then
        If rsPDOColab.RecordCount > 0 Then
            vNumPDO = rsPDOColab.Fields(3)
            If rsPDOColab.Fields(4) = "N" Or IsNull(rsPDOColab.Fields(4)) Then
                MsgBox "O PDO n�: " & Format(vNumPDO, "000000") & " esta em aberto para este " & rsPDOColab.Fields(5) & ". Aguarde tomada de decis�o", vbCritical, "Aten��o"
                rsPDOColab.Close
                Set rsPDOColab = Nothing
                Exit Function
            Else
                If Not IsNull(rsPDOColab.Fields(4)) Then
                    vStatusPDO = rsPDOColab.Fields(4)
                    vDecisao = rsPDOColab.Fields(6)
                End If
            End If
        End If
    End If
    rsPDOColab.Close
    Set rsPDOColab = Nothing
    
    If vStatusPDO <> "S" Then
        If Val(RemoveMask(Label41)) < MediaGlobal And Val(RemoveMask(Label41)) >= vAprovadoRest Then
            If vAdiRes = "N" Then
                If MsgBox("Pontua��o do colaborador est� abaixo da m�dia. Us�ario n�o privil�gios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                    gravaSolicitacao txtNovoColaborador(0), "colaborador", RemoveMask(Label41), "O colaborador est� participando do PS. P�rem, sua pontua��o est� abaixo da m�dia parametrizada no sistema para Aprova��o no cargo: " & Label5, NomUsu
                    MsgBox "Foi gerado o PDO n�: " & Format(vPDO, "000000") & ". Aguarde tomada de decis�o", vbInformation, "SGCH"
                End If
                'configControles
                Exit Function
            End If
        End If
        If Val(RemoveMask(Label41)) < vAprovadoRest Then
            If vAdiRep = "N" Then
                If MsgBox("Pontua��o do colaborador est� abaixo da m�dia. Us�ario n�o privil�gios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                    gravaSolicitacao txtNovoColaborador(0), "colaborador", RemoveMask(Label41), "O colaborador est� participando do PS. P�rem, sua pontua��o est� abaixo da m�dia parametrizada no sistema para Aprova��o com Restri��o no cargo: " & Label5, NomUsu
                    MsgBox "Foi gerado o PDO n�: " & Format(vPDO, "000000") & ". Aguarde tomada de decis�o", vbInformation, "SGCH"
                End If
                'configControles
                Exit Function
            End If
        End If
    Else
        If Trim(vDecisao) <> "Aprovado" Then
            MsgBox "O PDO n�: " & Format(vNumPDO, "000000") & " N�O FOI APROVADO ", vbCritical, "Aten��o"

            'Remover Numero de PDO da tabela de colaboradores
            SqlPDOColab = "Update tbColaboradores set autorizacao = Null Where a.codcoligada = '" & vCodcoligada & "' and cpf = '" & txtNovoColaborador(0) & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
            Exit Function
        Else
            'Remover Numero de PDO da tabela de colaboradores
            SqlPDOColab = "Update tbColaboradores set autorizacao = Null Where a.codcoligada = '" & vCodcoligada & "' and cpf = '" & txtNovoColaborador(0) & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
        End If
    End If
    avaliaAdmissao = True
End Function

Private Sub comporCombosTotvs()
    Dim X As Integer
    CompoeComboTotvs Combo(9), "PFUNCAO", "codigo", "nome"
End Sub

Private Sub comporControlesTotvs()
    On Error Resume Next
    Dim rsContrTotvs As New ADODB.Recordset
    Dim SqlContrTotvs As String
        
    SqlContrTotvs = "select * from tbColaboradoresIntTotvs where codcoligada = '" & vCodcoligada & "' and id = '" & Val(Label17) & "'"
    rsContrTotvs.Open SqlContrTotvs, cnBanco, adOpenKeyset, adLockReadOnly
    txtCons(8) = rsContrTotvs.Fields(10)
    txtCons_KeyDown 8, 13, 8
    rsContrTotvs.Close
    Set rsContrTotvs = Nothing
End Sub

Private Sub txtCons_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(8) <> "" Then CarregaComboTotvs "PFUNCAO", "CODIGO", txtCons(8).Text, Combo(9).Text, Index, "nome"
        End If
    End Select
End Sub

Private Sub Combo_Click(Index As Integer)
    Select Case Index
    Case 9
        AchaComboTotvs Combo(Index), "PFUNCAO", "CODIGO", Index, "nome"
    End Select
End Sub

