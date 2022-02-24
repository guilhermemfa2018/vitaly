VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "frmUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Coligada"
      Height          =   615
      Left            =   5280
      TabIndex        =   44
      Top             =   6720
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Usuário MULTIPLICADOR"
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informe os cursos/treinamentos"
      Enabled         =   0   'False
      Height          =   6015
      Left            =   5280
      TabIndex        =   36
      Top             =   600
      Width           =   5535
      Begin SGCH.chameleonButton chameleonButton2 
         Height          =   615
         Left            =   1440
         TabIndex        =   43
         Tag             =   "retirar treinamento"
         ToolTipText     =   "retirar treinamento"
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
         MICON           =   "frmUsuarios.frx":0CCA
         PICN            =   "frmUsuarios.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton chameleonButton1 
         Height          =   615
         Left            =   840
         TabIndex        =   42
         Tag             =   "Incluir Treinamento"
         ToolTipText     =   "Incluir Treinamento"
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
         MICON           =   "frmUsuarios.frx":19C0
         PICN            =   "frmUsuarios.frx":19DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   150
         TabIndex        =   41
         Top             =   3465
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   135
         TabIndex        =   40
         Top             =   345
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   39
         Top             =   3720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4048
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
         NumItems        =   0
      End
   End
   Begin SGCH.chameleonButton chameleonButton11 
      Height          =   615
      Left            =   720
      TabIndex        =   33
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   6720
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
      MICON           =   "frmUsuarios.frx":26B6
      PICN            =   "frmUsuarios.frx":26D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9720
      TabIndex        =   34
      Top             =   6720
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Conta "
      Height          =   2295
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   5055
      Begin VB.ComboBox cboCadastro 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   31
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox chkCadastro 
         Caption         =   "O usuário deve alterar Login e Senha no próximo Logon"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Grupo:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Login:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Senha:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Confirmar senha:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   12
      Top             =   3840
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   5
      Left            =   960
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskCadastro 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(##)####-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskCadastro 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(##)####-####"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboCadastro 
      Height          =   315
      Index           =   0
      ItemData        =   "frmUsuarios.frx":33AC
      Left            =   960
      List            =   "frmUsuarios.frx":3401
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin MSMask.MaskEdBox mskCadastro 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin SGCH.chameleonButton chameleonButton12 
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   6720
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
      MICON           =   "frmUsuarios.frx":3471
      PICN            =   "frmUsuarios.frx":348D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Ramal:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Celular:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Bairro:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblcadastro 
      Caption         =   "Nome:"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblcadastro 
      Caption         =   "Código:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsFuncionarios As New ADODB.Recordset
Private SqlFuncionarios As String

Private rsSalvar As New ADODB.Recordset
Private rsSenha As New ADODB.Recordset
Private Status As String

Private Sub chameleonButton1_Click()
    addRemLoteNota ListView1, ListView2
End Sub

Private Sub chameleonButton2_Click()
    addRemLoteNota ListView2, ListView1
End Sub

Private Sub chameleonButton11_Click()
    Unload Me
    Set frmUsuarios = Nothing
End Sub

Private Sub chameleonButton12_Click()
    If MsgBox("Deseja salvar os dados do usuário?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
        Bot_salvar
        gravaLog "Código usuário: " & mskCadastro(0), "Nome: " & txtCadastro(0), "Login: " & txtCadastro(7)
        Unload Me
        Set frmUsuarios = Nothing
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Frame2.Enabled = True
        chameleonButton1.UseGreyscale = False
        chameleonButton2.UseGreyscale = False
    Else
        Frame2.Enabled = False
        chameleonButton1.UseGreyscale = True
        chameleonButton2.UseGreyscale = True
    End If
End Sub

Private Sub Check3_Click()
    MarcaDesmarca ListView1
End Sub

Private Sub Check4_Click()
    MarcaDesmarca ListView2
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    listview_cabecalho
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        DesbloqueiaControles
    End If
    CompoeCombo1 cboCadastro(1), "tbgrupo", "codigo", "descricao"
    CompoeCombo2 Combo1, "tbDadosEmpresa", "codcoligada", "razaosocial"
    If txtCadastro(7) = "adm" Then
        txtCadastro(7).Enabled = False
        Combo1.Enabled = False
    Else
        txtCadastro(7).Enabled = True
        Combo1.Enabled = True
    End If
    
    lista_Treinamentos
    configControles
End Sub

Private Sub ResultPesq()
    SqlFuncionarios = "Select * from tbUsuarios, tbgrupo, tbsenha,tbdadosempresa where tbUsuarios.codigo = tbsenha.codigo and tbUsuarios.codgrupo = tbgrupo.codigo and tbUsuarios.codcoligada = tbdadosempresa.codcoligada and tbUsuarios.codigo = '" & Val(varGlobal) & "' order by tbUsuarios.codigo"
    rsFuncionarios.Open SqlFuncionarios, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFuncionarios.RecordCount > 0 Then
        CompoeControles
    Else
        MsgBox "Usuário não encontrado"
    End If
    rsFuncionarios.Close
    Set rsFuncionarios = Nothing
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DesbloqueiaControles
    For X = 0 To mskCadastro.Count - 1
        Me.mskCadastro(X).PromptInclude = False
        mskCadastro(X) = ""
        Me.mskCadastro(X).PromptInclude = True
    Next
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X) = ""
    Next
    For X = 0 To cboCadastro.Count - 1
        cboCadastro(X) = ""
    Next
    mskCadastro(0).Text = Format(GeraCodigo, "000000") & ""
End Sub

Private Sub CompoeControles()
    mskCadastro(0).Enabled = False
    mskCadastro(0).Text = Format(rsFuncionarios.Fields(0), "000000") & ""
    mskCadastro(0).PromptInclude = True
    mskCadastro(1).PromptInclude = False
    mskCadastro(1).Text = rsFuncionarios.Fields(7) & ""
    mskCadastro(1).PromptInclude = True
    mskCadastro(2).PromptInclude = False
    mskCadastro(2).Text = rsFuncionarios.Fields(8) & ""
    mskCadastro(2).PromptInclude = True
    txtCadastro(0).Text = rsFuncionarios.Fields(1)
    If rsFuncionarios.Fields(2) <> "Null" Then txtCadastro(1).Text = rsFuncionarios.Fields(2)
    If rsFuncionarios.Fields(3) <> "Null" Then txtCadastro(2).Text = rsFuncionarios.Fields(3)
    If rsFuncionarios.Fields(3) <> "Null" Then txtCadastro(3).Text = rsFuncionarios.Fields(4)
    If rsFuncionarios.Fields(4) <> "Null" Then txtCadastro(4).Text = rsFuncionarios.Fields(5)
    If rsFuncionarios.Fields(5) <> "Null" Then txtCadastro(5).Text = rsFuncionarios.Fields(9)
    If rsFuncionarios.Fields(6) <> "Null" Then txtCadastro(6).Text = rsFuncionarios.Fields(10)
    
    txtCadastro(7).Text = rsFuncionarios.Fields(20)
    txtCadastro(8).Text = rsFuncionarios.Fields(21)
    txtCadastro(9).Text = rsFuncionarios.Fields(21)
    cboCadastro(0).Text = rsFuncionarios.Fields(6) & ""
    cboCadastro(1).Text = Format(rsFuncionarios.Fields(16), "000000") & " - " & rsFuncionarios.Fields(17) & ""
    If rsFuncionarios.Fields(12) = 1 Then
        chkCadastro(0).Value = 1
    Else
        chkCadastro(0).Value = 0
    End If
    If rsFuncionarios.Fields(13) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If rsFuncionarios.Fields(14) = "S" Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
    Combo1.Text = rsFuncionarios.Fields(24)
    BloqueiaControles
End Sub

Private Sub Bot_salvar()
'On Error GoTo TrataErro
    If txtCadastro(8).Text <> txtCadastro(9).Text Then
        MsgBox "A nova senha e a senha de confirmação devem ser iguais. Digite-as novamente", vbCritical, "SGCH"
        Exit Sub
    End If
    
    If cboCadastro(1).Text = "" Then
        MsgBox "Selecione o Grupo ao qual pertence o usuário", vbCritical, "SGCH"
        Exit Sub
    End If
    If Combo1.Text = "" Then
        MsgBox "Selecione a COLIGADA a qual pertence o usuário", vbCritical, "SGCH"
        Exit Sub
    End If
    Dim SqlSalvar As String
    Dim X As Integer, Y As Integer
    Dim rsSalvTrei As New ADODB.Recordset
    Dim SqlSalvTrei As String
    Dim rsColigada As New ADODB.Recordset
    Dim SqlColigada As String
    
    cnBanco.BeginTrans
    SqlColigada = "Select codcoligada from tbDadosEmpresa where tbDadosEmpresa.razaosocial = '" & Combo1 & "'"
    rsColigada.Open SqlColigada, cnBanco, adOpenKeyset, adLockReadOnly
    If rsColigada.RecordCount > 0 Then
        vCodcoligada = rsColigada.Fields(0)
    End If
    rsColigada.Close
    Set rsColigada = Nothing
    
    SqlSalvar = "Select * from tbUsuarios where tbUsuarios.codigo = '" & Val(Me.mskCadastro(0)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    mskCadastro(0).PromptInclude = False
    If txtCadastro(0).Text <> "" Then
        mskCadastro(0).PromptInclude = False
        If mskCadastro(0).Text <> "" Then
            'cnBanco.BeginTrans
            If rsSalvar.RecordCount = 0 Then
                rsSalvar.AddNew
                rsSalvar.Fields(0) = mskCadastro(0).ClipText
                rsSalvar.Fields(7) = mskCadastro(1).ClipText
                rsSalvar.Fields(8) = mskCadastro(2).ClipText
                rsSalvar.Fields(1) = txtCadastro(0).Text
                rsSalvar.Fields(2) = txtCadastro(1).Text
                rsSalvar.Fields(3) = txtCadastro(2).Text
                rsSalvar.Fields(4) = txtCadastro(3).Text
                rsSalvar.Fields(5) = txtCadastro(4).Text
                rsSalvar.Fields(9) = txtCadastro(5).Text
                rsSalvar.Fields(10) = txtCadastro(6).Text
                rsSalvar.Fields(6) = cboCadastro(0).Text
                
                rsSalvar.Fields(11) = Val(Left(cboCadastro(1).Text, 6))
                
                If chkCadastro(0).Value = 0 Then
                    rsSalvar.Fields(12) = 0
                ElseIf chkCadastro(0).Value = 1 Then
                    rsSalvar.Fields(12) = 1
                End If
                
                If Check1.Value = 0 Then
                    rsSalvar.Fields(13) = "N"
                Else
                    rsSalvar.Fields(13) = "S"
                End If
                If Check2.Value = 0 Then
                    rsSalvar.Fields(14) = "N"
                Else
                    rsSalvar.Fields(14) = "S"
                End If
                rsSalvar.Fields(15) = vCodcoligada 'Codigo da coligada
                rsSalvar.Update

                SqlSalvar = "Select * from tbsenha where tbsenha.codigo = '" & Val(Me.mskCadastro(0)) & "'"
                rsSenha.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
                
                rsSenha.AddNew
                rsSenha.Fields(2) = mskCadastro(0).ClipText
                rsSenha.Fields(1) = txtCadastro(8)
                rsSenha.Fields(0) = txtCadastro(7).Text
                rsSenha.Fields(3) = vCodcoligada 'Codigo da coligada
                MsgBox "Inclusão realizada com sucesso!"
                LimpaControles
            Else
                rsSalvar.Fields(0) = mskCadastro(0).ClipText
                rsSalvar.Fields(7) = mskCadastro(1).ClipText
                rsSalvar.Fields(8) = mskCadastro(2).ClipText
                rsSalvar.Fields(1) = txtCadastro(0).Text
                rsSalvar.Fields(2) = txtCadastro(1).Text
                rsSalvar.Fields(3) = txtCadastro(2).Text
                rsSalvar.Fields(4) = txtCadastro(3).Text
                rsSalvar.Fields(5) = txtCadastro(4).Text
                rsSalvar.Fields(9) = txtCadastro(5).Text
                rsSalvar.Fields(10) = txtCadastro(6).Text
                rsSalvar.Fields(6) = cboCadastro(0).Text
                
                rsSalvar.Fields(11) = Val(Left(cboCadastro(1).Text, 6))
                If chkCadastro(0).Value = 0 Then
                    rsSalvar.Fields(12) = 0
                ElseIf chkCadastro(0).Value = 1 Then
                    rsSalvar.Fields(12) = 1
                End If
                If Check1.Value = 0 Then
                    rsSalvar.Fields(13) = "N"
                Else
                    rsSalvar.Fields(13) = "S"
                End If
                If Check2.Value = 0 Then
                    rsSalvar.Fields(14) = "N"
                Else
                    rsSalvar.Fields(14) = "S"
                End If
                
                rsSalvar.Fields(15) = vCodcoligada 'Codigo da coligada
                rsSalvar.Update
                
                SqlSalvar = "Select * from tbsenha where tbsenha.codigo = '" & Val(Me.mskCadastro(0)) & "'"
                rsSenha.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
                
                rsSenha.Fields(1) = txtCadastro(8).Text
                rsSenha.Fields(0) = txtCadastro(7).Text
                rsSenha.Fields(3) = vCodcoligada  'Codigo da coligada
                
'- Grava trinamentos do MULTIPLICADOR
                If Check2.Value = 1 And ListView2.ListItems.Count > 0 Then
                    SqlSalvTrei = "Delete from tbUsuMultiplic where tbUsuMultiplic.codusuario= " & Val(Me.mskCadastro(0))
                    rsSalvTrei.Open SqlSalvTrei, cnBanco
    
                    SqlSalvTrei = "Select * from tbUsuMultiplic where tbUsuMultiplic.codusuario= " & Val(Me.mskCadastro(0))
                    rsSalvTrei.Open SqlSalvTrei, cnBanco, adOpenKeyset, adLockOptimistic
    
                    Y = ListView2.ListItems.Count
                    For X = 1 To Y
                        rsSalvTrei.AddNew
                        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
                        rsSalvTrei.Fields(0) = Val(Me.mskCadastro(0))
                        rsSalvTrei.Fields(1) = ListView2.ListItems.Item(X)
                        rsSalvTrei.Fields(2) = vCodcoligada 'Codigo da coligada
                    Next
                    rsSalvTrei.Update
                    rsSalvTrei.Close
                    Set rsSalvTrei = Nothing
                Else
                    SqlSalvTrei = "Delete from tbUsuMultiplic where tbUsuMultiplic.codusuario= " & Val(Me.mskCadastro(0))
                    rsSalvTrei.Open SqlSalvTrei, cnBanco
                End If
'----------------------------------
            
            MsgBox "Alteração realizada com sucesso!"
            End If
            'rsSalvar.Update
            rsSenha.Update
            cnBanco.CommitTrans
            rsSalvar.Close
            Set rsSalvar = Nothing
            rsSenha.Close
            Set rsSenha = Nothing
        Else
            MsgBox "Favor Preencher o campo código!"
        End If
    Else
        MsgBox "Favor Preencher o campo Nome!"
    End If
    AtualizaListview
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(mskCadastro(0), "000000"))
        ItemLst.SubItems(1) = txtCadastro(0).Text
        ItemLst.SubItems(2) = Mid$(cboCadastro(1).Text, 10, 20)
        If Check1.Value = 0 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadastro(0).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = Mid$(cboCadastro(1).Text, 10, 20)
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

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirUsuario
    SqlGera = "Select top 1 * from tbUsuarios order by codigo Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFuncionarios.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    mskCadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharUsuario
End Function

Private Sub AbrirUsuario()
    SqlFuncionarios = "Select * from tbUsuarios Order by codigo"
    rsFuncionarios.Open SqlFuncionarios, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharUsuario()
    rsFuncionarios.Close
    Set rsFuncionarios = Nothing
End Sub

Private Function DesbloqueiaControles()
    Dim X As Integer
    
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X).Enabled = True
    Next
    For X = 0 To mskCadastro.Count - 1
        mskCadastro(X).Enabled = True
    Next
    For X = 0 To cboCadastro.Count - 1
        cboCadastro(X).Enabled = True
    Next
    chkCadastro(0).Enabled = True
    mskCadastro(0).Enabled = False
    If Check2.Value = 1 Then
        Frame2.Enabled = True
        chameleonButton1.UseGreyscale = False
        chameleonButton2.UseGreyscale = False
    Else
        Frame2.Enabled = False
        chameleonButton1.UseGreyscale = True
        chameleonButton2.UseGreyscale = True
    End If
End Function

Private Function BloqueiaControles()
    Dim X As Integer
    For X = 0 To txtCadastro.Count - 1
        txtCadastro(X).Enabled = False
    Next
    For X = 0 To mskCadastro.Count - 1
        mskCadastro(X).Enabled = False
    Next
    For X = 0 To cboCadastro.Count - 1
        cboCadastro(X).Enabled = False
    Next
    chkCadastro(0).Enabled = False
    
    If Check2.Value = 1 Then
        Frame2.Enabled = True
        chameleonButton1.UseGreyscale = False
        chameleonButton2.UseGreyscale = False
    Else
        Frame2.Enabled = False
        chameleonButton1.UseGreyscale = True
        chameleonButton2.UseGreyscale = True
    End If
End Function

Private Sub configControles()
    If vSal = "N" Then
        chameleonButton12.UseGreyscale = True
        chameleonButton12.DragMode = 1
        chameleonButton12.SpecialEffect = cbEngraved
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Nome do treinamento", ListView1.Width / 1.3
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview

    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Código", ListView2.Width / 7
    ListView2.ColumnHeaders.Add , , "Nome do treinamento", ListView2.Width / 1.3
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview

End Sub

Private Sub lista_Treinamentos()
    Dim rsListaTreinamento As New ADODB.Recordset
    Dim sqlListaTreinamento As String
    Dim ItemLst As ListItem
    Dim ItemLst2 As ListItem
    Me.mskCadastro(0).PromptInclude = False
    sqlListaTreinamento = "select a.codtreinamento,a.nometreinamento,b.codtreinamento from tbtreinamentos as a left join tbUsuMultiplic as b on b.codtreinamento = a.codtreinamento and b.codusuario = '" & Val(mskCadastro(0)) & "' where a.ativo = 'S' order by a.codtreinamento"
    rsListaTreinamento.Open sqlListaTreinamento, cnBanco, adOpenKeyset, adLockReadOnly
    Me.mskCadastro(0).PromptInclude = True
    While Not rsListaTreinamento.EOF
        If IsNull(rsListaTreinamento.Fields(2)) Then
            Set ItemLst = ListView1.ListItems.Add(, , rsListaTreinamento.Fields(0))
            ItemLst.SubItems(1) = "" & rsListaTreinamento.Fields(1)
        Else
            Set ItemLst2 = ListView2.ListItems.Add(, , rsListaTreinamento.Fields(0))
            ItemLst2.SubItems(1) = "" & rsListaTreinamento.Fields(1)
        End If
        rsListaTreinamento.MoveNext
    Wend
End Sub

Private Sub MarcaDesmarca(LV As ListView)
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems(X).Selected = True
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        Else
            LV.ListItems.Item(X).Checked = True
        End If
    Next
End Sub

Private Sub addRemLoteNota(lvOrigem As ListView, lvDestino As ListView)
    Dim X As Integer, Y As Integer
    Dim ItemLst As ListItem
    Y = lvOrigem.ListItems.Count
    For X = 1 To Y
        If Y < X Then
            Exit Sub
        End If
        lvOrigem.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If lvOrigem.ListItems(X).Checked = True Then
            Set ItemLst = lvDestino.ListItems.Add(, , lvOrigem.ListItems(X)) ' Copiar o primeiro item e criar o CheckBox
            ItemLst.SubItems(1) = "" & lvOrigem.SelectedItem.ListSubItems.Item(1) 'Copia o coluna que vc desejar do item selecionado
            lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
            Y = Y - 1
            X = X - 1
        End If
    Next
    'Ordena listview para exibir na tela
    lvDestino.Sorted = True
    lvDestino.SortKey = 0
    lvDestino.SortOrder = lvwAscending
    lvDestino.Refresh
End Sub
