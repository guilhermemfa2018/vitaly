VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
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
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0CCA
      TabIndex        =   42
      Top             =   3960
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0D34
      TabIndex        =   41
      Top             =   3600
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0D9E
      TabIndex        =   40
      Top             =   3240
      Width           =   735
   End
   Begin MSMask.MaskEdBox mskCadastro 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(##)####-####"
      PromptChar      =   "_"
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0E0C
      TabIndex        =   39
      Top             =   2880
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0E7C
      TabIndex        =   38
      Top             =   2520
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0EE8
      TabIndex        =   37
      Top             =   2160
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0F54
      TabIndex        =   36
      Top             =   1800
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":0FC0
      TabIndex        =   35
      Top             =   1440
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":1026
      TabIndex        =   34
      Top             =   1080
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":1096
      TabIndex        =   33
      Top             =   720
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmUsuarios.frx":10FE
      TabIndex        =   32
      Top             =   360
      Width           =   615
   End
   Begin ZEUS.chameleonButton chameleonButton11 
      Height          =   615
      Left            =   720
      TabIndex        =   18
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
      MICON           =   "frmUsuarios.frx":116A
      PICN            =   "frmUsuarios.frx":1186
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton chameleonButton12 
      Height          =   615
      Left            =   120
      TabIndex        =   17
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
      MICON           =   "frmUsuarios.frx":1E60
      PICN            =   "frmUsuarios.frx":1E7C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      Caption         =   "Coligada"
      Height          =   615
      Left            =   5280
      TabIndex        =   29
      Top             =   6720
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Autorizado a finalizar apropriação de colaboradores"
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   240
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informe os cursos/treinamentos"
      Enabled         =   0   'False
      Height          =   6015
      Left            =   5280
      TabIndex        =   21
      Top             =   600
      Width           =   5535
      Begin ZEUS.chameleonButton chameleonButton2 
         Height          =   615
         Left            =   1440
         TabIndex        =   28
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
         MICON           =   "frmUsuarios.frx":2B56
         PICN            =   "frmUsuarios.frx":2B72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton chameleonButton1 
         Height          =   615
         Left            =   840
         TabIndex        =   27
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
         MICON           =   "frmUsuarios.frx":384C
         PICN            =   "frmUsuarios.frx":3868
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
         TabIndex        =   26
         Top             =   3465
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   345
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   24
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
         TabIndex        =   23
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
   Begin VB.Frame Frame1 
      Caption         =   " Conta "
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   5055
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":4542
         TabIndex        =   31
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1080
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":45AC
         TabIndex        =   45
         Top             =   1200
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":462A
         TabIndex        =   44
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":4694
         TabIndex        =   43
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboCadastro 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox chkCadastro 
         Caption         =   "O usuário deve alterar Login e Senha no próximo Logon"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   5
      Left            =   960
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cboCadastro 
      Height          =   315
      Index           =   0
      ItemData        =   "frmUsuarios.frx":46FE
      Left            =   960
      List            =   "frmUsuarios.frx":4753
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtCadastro 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   1
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
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   6720
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
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
    mobjMsg.Abrir "Deseja salvar os dados do usuário?", YesNo, pergunta, "ZEUS"
    If Tp = 1 Then
        Bot_salvar
        gravaLog "Código usuário: " & mskCadastro(0), "Nome: " & txtcadastro(0), "Login: " & txtcadastro(7)
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
    If txtcadastro(7) = "adm" Then
        txtcadastro(7).Enabled = False
        Combo1.Enabled = False
    Else
        txtcadastro(7).Enabled = True
        Combo1.Enabled = True
    End If
    
'    lista_Treinamentos
    configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    SqlFuncionarios = "Select * from tbUsuarios, tbgrupo, tbsenha,tbdadosempresa where tbUsuarios.codigo = tbsenha.codigo and tbUsuarios.codgrupo = tbgrupo.codigo and tbUsuarios.codcoligada = tbdadosempresa.codcoligada and tbUsuarios.codigo = '" & Val(varGlobal) & "' order by tbUsuarios.codigo"
    rsFuncionarios.Open SqlFuncionarios, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFuncionarios.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Usuário não encontrado", Ok, critico, "Atenção"
    End If
    rsFuncionarios.Close
    Set rsFuncionarios = Nothing
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

Private Sub LimpaControles()
    Dim x As Integer
    DesbloqueiaControles
    For x = 0 To mskCadastro.Count - 1
        Me.mskCadastro(x).PromptInclude = False
        mskCadastro(x) = ""
        Me.mskCadastro(x).PromptInclude = True
    Next
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x) = ""
    Next
    For x = 0 To cboCadastro.Count - 1
        cboCadastro(x) = ""
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
    txtcadastro(0).Text = rsFuncionarios.Fields(1)
    If rsFuncionarios.Fields(2) <> "Null" Then txtcadastro(1).Text = rsFuncionarios.Fields(2)
    If rsFuncionarios.Fields(3) <> "Null" Then txtcadastro(2).Text = rsFuncionarios.Fields(3)
    If rsFuncionarios.Fields(3) <> "Null" Then txtcadastro(3).Text = rsFuncionarios.Fields(4)
    If rsFuncionarios.Fields(4) <> "Null" Then txtcadastro(4).Text = rsFuncionarios.Fields(5)
    If rsFuncionarios.Fields(5) <> "Null" Then txtcadastro(5).Text = rsFuncionarios.Fields(9)
    If rsFuncionarios.Fields(6) <> "Null" Then txtcadastro(6).Text = rsFuncionarios.Fields(10)
    
    txtcadastro(7).Text = rsFuncionarios.Fields(20)
    txtcadastro(8).Text = rsFuncionarios.Fields(21)
    txtcadastro(9).Text = rsFuncionarios.Fields(21)
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
On Error GoTo Err
    If txtcadastro(8).Text <> txtcadastro(9).Text Then
        mobjMsg.Abrir "A nova senha e a senha de confirmação devem ser iguais. Digite-as novamente", Ok, critico, "Atenção"
        Exit Sub
    End If
    
    If cboCadastro(1).Text = "" Then
        mobjMsg.Abrir "Selecione o Grupo ao qual pertence o usuário", Ok, critico, "Atenção"
        Exit Sub
    End If

    Dim SqlSalvar As String
    Dim x As Integer, y As Integer
    Dim rsSalvTrei As New ADODB.Recordset
    Dim SqlSalvTrei As String
    Dim rsColigada As New ADODB.Recordset
    Dim SqlColigada As String
    
10  cnBanco.BeginTrans
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
    If txtcadastro(0).Text <> "" Then
        mskCadastro(0).PromptInclude = False
        If mskCadastro(0).Text <> "" Then
            'cnBanco.BeginTrans
            If rsSalvar.RecordCount = 0 Then
                rsSalvar.AddNew
                rsSalvar.Fields(0) = mskCadastro(0).ClipText
                rsSalvar.Fields(7) = mskCadastro(1).ClipText
                rsSalvar.Fields(8) = mskCadastro(2).ClipText
                rsSalvar.Fields(1) = txtcadastro(0).Text
                rsSalvar.Fields(2) = txtcadastro(1).Text
                rsSalvar.Fields(3) = txtcadastro(2).Text
                rsSalvar.Fields(4) = txtcadastro(3).Text
                rsSalvar.Fields(5) = txtcadastro(4).Text
                rsSalvar.Fields(9) = txtcadastro(5).Text
                rsSalvar.Fields(10) = txtcadastro(6).Text
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
                rsSenha.Fields(1) = txtcadastro(8)
                rsSenha.Fields(0) = txtcadastro(7).Text
                rsSenha.Fields(3) = vCodcoligada 'Codigo da coligada
                mobjMsg.Abrir "Inclusão realizada com sucesso!", Ok, informacao, "Atenção"
                LimpaControles
            Else
                rsSalvar.Fields(0) = mskCadastro(0).ClipText
                rsSalvar.Fields(7) = mskCadastro(1).ClipText
                rsSalvar.Fields(8) = mskCadastro(2).ClipText
                rsSalvar.Fields(1) = txtcadastro(0).Text
                rsSalvar.Fields(2) = txtcadastro(1).Text
                rsSalvar.Fields(3) = txtcadastro(2).Text
                rsSalvar.Fields(4) = txtcadastro(3).Text
                rsSalvar.Fields(5) = txtcadastro(4).Text
                rsSalvar.Fields(9) = txtcadastro(5).Text
                rsSalvar.Fields(10) = txtcadastro(6).Text
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
                
                'rsSalvar.Fields(15) = vCodcoligada 'Codigo da coligada
                rsSalvar.Update
                
                SqlSalvar = "Select * from tbsenha where tbsenha.codigo = '" & Val(Me.mskCadastro(0)) & "'"
                rsSenha.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
                
                rsSenha.Fields(1) = txtcadastro(8).Text
                rsSenha.Fields(0) = txtcadastro(7).Text
                rsSenha.Fields(3) = vCodcoligada  'Codigo da coligada
                
'----------------------------------
            mobjMsg.Abrir "Alteração realizada com sucesso!", Ok, informacao, "Atenção"
            End If
            'rsSalvar.Update
            rsSenha.Update
            cnBanco.CommitTrans
            rsSalvar.Close
            Set rsSalvar = Nothing
            rsSenha.Close
            Set rsSenha = Nothing
        Else
            mobjMsg.Abrir "Favor Preencher o campo código", Ok, critico, "Atenção"
        End If
    Else
        mobjMsg.Abrir "Favor Preencher o campo Nome", Ok, critico, "Atenção"
    End If
    AtualizaListview
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim y As Integer, x As Integer
    y = MeuLV.ListView1.ListItems.Count
    For x = 1 To y
        If MeuLV.ListView1.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(mskCadastro(0), "000000"))
        ItemLst.SubItems(1) = txtcadastro(0).Text
        ItemLst.SubItems(2) = Mid$(cboCadastro(1).Text, 10, 20)
        If Check1.Value = 0 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtcadastro(0).Text
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
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Function GeraCodigo()
On Error GoTo Err
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

Private Sub AbrirUsuario()
On Error GoTo Err
    SqlFuncionarios = "Select * from tbUsuarios Order by codigo"
    rsFuncionarios.Open SqlFuncionarios, cnBanco, adOpenKeyset, adLockOptimistic
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

Private Sub FecharUsuario()
    rsFuncionarios.Close
    Set rsFuncionarios = Nothing
End Sub

Private Function DesbloqueiaControles()
    Dim x As Integer
    
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = True
    Next
    For x = 0 To mskCadastro.Count - 1
        mskCadastro(x).Enabled = True
    Next
    For x = 0 To cboCadastro.Count - 1
        cboCadastro(x).Enabled = True
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
    Dim x As Integer
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = False
    Next
    For x = 0 To mskCadastro.Count - 1
        mskCadastro(x).Enabled = False
    Next
    For x = 0 To cboCadastro.Count - 1
        cboCadastro(x).Enabled = False
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
On Error GoTo Err
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


Private Sub addRemLoteNota(lvOrigem As Listview, lvDestino As Listview)
    Dim x As Integer, y As Integer
    Dim ItemLst As ListItem
    y = lvOrigem.ListItems.Count
    For x = 1 To y
        If y < x Then
            Exit Sub
        End If
        lvOrigem.ListItems.Item(x).Selected = True 'Passar a selecao para o próximo item
        If lvOrigem.ListItems(x).Checked = True Then
            Set ItemLst = lvDestino.ListItems.Add(, , lvOrigem.ListItems(x)) ' Copiar o primeiro item e criar o CheckBox
            ItemLst.SubItems(1) = "" & lvOrigem.SelectedItem.ListSubItems.Item(1) 'Copia o coluna que vc desejar do item selecionado
            lvOrigem.ListItems.Remove (x) ' Remove item selecionado do ListView1
            y = y - 1
            x = x - 1
        End If
    Next
    'Ordena listview para exibir na tela
    lvDestino.Sorted = True
    lvDestino.SortKey = 0
    lvDestino.SortOrder = lvwAscending
    lvDestino.Refresh
End Sub

