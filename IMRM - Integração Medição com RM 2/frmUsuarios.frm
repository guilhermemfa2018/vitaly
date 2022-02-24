VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Informe os locais de estoque que o usuário acessa"
      Height          =   6495
      Left            =   7200
      TabIndex        =   33
      Top             =   240
      Width           =   5535
      Begin IMRM.chameleonButton chameleonButton2 
         Height          =   615
         Left            =   1200
         TabIndex        =   39
         Tag             =   "Retirar local de estoque"
         ToolTipText     =   "Retirar local de estoque"
         Top             =   3360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUsuarios.frx":0CCA
         PICN            =   "frmUsuarios.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton chameleonButton1 
         Height          =   615
         Left            =   600
         TabIndex        =   38
         Tag             =   "Incluir local de estoque"
         ToolTipText     =   "Incluir local de estoque"
         Top             =   3360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUsuarios.frx":19C0
         PICN            =   "frmUsuarios.frx":19DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   135
         TabIndex        =   35
         Top             =   345
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   3705
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   120
         TabIndex        =   36
         Top             =   4080
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4683
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dados do Usuário (Totvs)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   26
      Top             =   1920
      Width           =   6975
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   4200
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame6 
         Caption         =   "Coligada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   6735
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   6495
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":26B6
         TabIndex        =   30
         Top             =   1800
         Width           =   6735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   29
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmUsuarios.frx":2812
         TabIndex        =   28
         Top             =   240
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":2874
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do usuário (Ferramentaria)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   6735
      End
      Begin MSMask.MaskEdBox mskCadastro 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmUsuarios.frx":28DA
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":293C
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":29A2
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Coligada"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   21
      Top             =   6840
      Width           =   4335
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4095
      End
   End
   Begin IMRM.chameleonButton chameleonButton11 
      Height          =   615
      Left            =   720
      TabIndex        =   15
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
      MICON           =   "frmUsuarios.frx":2A06
      PICN            =   "frmUsuarios.frx":2A22
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin IMRM.chameleonButton chameleonButton12 
      Height          =   615
      Left            =   120
      TabIndex        =   14
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
      MICON           =   "frmUsuarios.frx":36FC
      PICN            =   "frmUsuarios.frx":3718
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
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11640
      TabIndex        =   16
      Top             =   6840
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Conta "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   6975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":43F2
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":4456
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":44CE
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmUsuarios.frx":4532
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   5415
      End
      Begin VB.CheckBox chkCadastro 
         Caption         =   "O usuário deve alterar Login e Senha no próximo Logon"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1695
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

Private Sub chameleonButton11_Click()
    Unload Me
    Set frmUsuarios = Nothing
End Sub

Private Sub chameleonButton12_Click()
    mobjMsg.Abrir "Deseja salvar os dados do usuário?", YesNo, pergunta, "IMRM"
    If Tp = 1 Then
        If Bot_salvar = True Then
        'gravaLog "Código usuário: " & mskCadastro(0), "Nome: " & txtCadastro(0), "Login: " & txtCadastro(7)
            Unload Me
            Set frmUsuarios = Nothing
        End If
    End If
End Sub

Private Sub chameleonButton2_Click()
    addRemLoteNota ListView2, ListView1
End Sub

Private Sub Check3_Click()
    MarcaDesmarcaTodos ListView1
End Sub

Private Sub Check4_Click()
    MarcaDesmarcaTodos ListView2
End Sub

Private Sub Form_Activate()
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    chamaSQL "select a.CODLOC,a.NOME from " & vBancoSAP & ".dbo.tloc as a where a.CODCOLIGADA = 1 and a.codfilial = 1 and a.INATIVO = 0"
    Compoe_Listview ListView1, Sqlp, "00"

'    chamaSQL "select a.CODLOC,a.NOME from tblocalestoque as a where a.codigo = " & Val(varGlobal)
'    Compoe_Listview ListView2, Sqlp, "00"

End Sub

'Private Sub Check2_Click()
'    If Check2.Value = 1 Then
'        Combo2.Enabled = True
'    Else
'        Combo2.Enabled = False
'    End If
'End Sub

Private Sub Form_Load()
    
    listview_cabecalho
    
    Status = Pesquisa
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        DesbloqueiaControles
    End If
    CompoeCombo1 cboCadastro(1), "tbgrupo", "codigo", "descricao"
    CompoeCombo2 Combo1, "tbDadosEmpresa", "codcoligada", "razaosocial"
    CompoeCombo1 Combo2, vBancoSAP & ".dbo.gcoligada", "codcoligada", "nomefantasia"
    If txtCadastro(7) = "adm" Then
        txtCadastro(7).Enabled = False
        Combo1.Enabled = False
    Else
        txtCadastro(7).Enabled = True
        Combo1.Enabled = True
    End If
    
    configControles
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
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Nome Local Estoque", ListView1.Width / 1.3
    
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 7
    ListView2.ColumnHeaders.Add , , "Nome Local Estoque", ListView2.Width / 1.3
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ResultPesq()
    SqlFuncionarios = "Select a.codigo,a.nome,a.codven,a.nomeven,a.email,a.codgrupo,a.altlogin,a.ativo,a.codcoligada,b.codigo,b.descricao,b.ativo,c.codcoligada,c.usuario,c.senha,c.codigo,d.codcoligada,d.razaosocial,d.endereco,d.bairro,d.cidade,d.uf,d.cep,d.email,d.site,d.telefone,d.fax,d.cnpj,d.ie,d.logo,d.codcoligada,d.ativa,a.codusuarioTOTVS,a.codcoligadatotvs from tbUsuarios as a, tbgrupo as b, tbsenha as c,tbdadosempresa as d where a.codigo = c.codigo and a.codgrupo = b.codigo and a.codcoligada = d.codcoligada and a.codigo = '" & Val(varGlobal) & "' order by a.codigo"
    rsFuncionarios.Open SqlFuncionarios, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFuncionarios.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Usuário não encontrado", Ok, critico, "Atenção"
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
    txtCadastro(0) = ""
    txtCadastro(1) = ""
    txtCadastro(2) = ""
    txtCadastro(6) = ""
    txtCadastro(7) = ""
    txtCadastro(8) = ""
    txtCadastro(9) = ""
    cboCadastro(1) = ""
    mskCadastro(0).Text = Format(GeraCodigo, "000000") & ""
End Sub

Private Sub CompoeControles()
    mskCadastro(0).Enabled = False
    mskCadastro(0).Text = Format(rsFuncionarios.Fields(0), "000000") & "" 'Código do usuário na IMRM
    mskCadastro(0).PromptInclude = True
    txtCadastro(0).Text = rsFuncionarios.Fields(1) 'Nome do usuário na IMRM
    If rsFuncionarios.Fields(6) <> "Null" Then txtCadastro(6).Text = rsFuncionarios.Fields(4) 'Email do usuário na IMRM
    If rsFuncionarios.Fields(2) <> "Null" Then txtCadastro(1).Text = rsFuncionarios.Fields(2) 'Código do usuário na TOTVS
    If rsFuncionarios.Fields(3) <> "Null" Then txtCadastro(2).Text = rsFuncionarios.Fields(3) 'Nome do usuário na TOTVS
    txtCadastro(7).Text = rsFuncionarios.Fields(13)
    txtCadastro(8).Text = rsFuncionarios.Fields(14)
    txtCadastro(9).Text = rsFuncionarios.Fields(14)
    cboCadastro(1).Text = Format(rsFuncionarios.Fields(9), "000000") & " - " & rsFuncionarios.Fields(10) & ""
    If rsFuncionarios.Fields(6) = 1 Then
        chkCadastro(0).Value = 1
    Else
        chkCadastro(0).Value = 0
    End If
    If rsFuncionarios.Fields(7) = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If Not IsNull(rsFuncionarios.Fields(32)) Then txtCadastro(3).Text = rsFuncionarios.Fields(32)
    Combo1.Text = rsFuncionarios.Fields(17)
    If Not IsNull(rsFuncionarios.Fields(33)) Then Combo2.Text = rsFuncionarios.Fields(33)
    BloqueiaControles
End Sub

Private Function Bot_salvar()
'On Error GoTo TrataErro
    Bot_salvar = True
    If txtCadastro(8).Text <> txtCadastro(9).Text Then
        mobjMsg.Abrir "A nova senha e a senha de confirmação devem ser iguais. Digite-as novamente", Ok, critico, "Atenção"
        Bot_salvar = False
        Exit Function
    End If
    If cboCadastro(1).Text = "" Then
        mobjMsg.Abrir "Selecione o Grupo ao qual pertence o usuário", Ok, critico, "Atenção"
        Bot_salvar = False
        Exit Function
    End If
    If txtCadastro(1).Text = "" Then
        mobjMsg.Abrir "Vincule o usuário a um usuário TOTVS", Ok, critico, "Atenção"
        Bot_salvar = False
        Exit Function
    End If

    If ListView2.ListItems.Count = 0 Then
        mobjMsg.Abrir "Deve ser selecionado 1 local de estoque", Ok, critico, "Atenção"
        Bot_salvar = False
        Exit Function
    End If

    Dim SqlSalvar As String
    Dim X As Integer, Y As Integer
    Dim rsSalvTrei As New ADODB.Recordset
    Dim SqlSalvTrei As String
    Dim rsColigada As New ADODB.Recordset
    Dim SqlColigada As String
    
    cnBanco.BeginTrans
    AtualizaListview
    
'    If ListView2.ListItems.Count > 0 Then gravaLocalEstoque

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
                rsSalvar.Fields(0) = mskCadastro(0).ClipText 'Código do usuário da IMRM
                rsSalvar.Fields(1) = txtCadastro(0).Text 'Nome do usuário da IMRM
                rsSalvar.Fields(4) = txtCadastro(6).Text 'Email do usuário da IMRM
                rsSalvar.Fields(2) = txtCadastro(1).Text 'codven na TOTVS
                rsSalvar.Fields(3) = txtCadastro(2).Text 'Nome do usuário na TOTVS
                rsSalvar.Fields(9) = Combo2.Text 'Codigo+nome Coligada TOTVS
                
                rsSalvar.Fields(5) = Val(Left(cboCadastro(1).Text, 6)) 'Grupo aoqual pertence o usuário da IMRM
                
                If chkCadastro(0).Value = 0 Then 'Informa se o usuário irá alterar a senha no próximo login
                    rsSalvar.Fields(6) = 0
                ElseIf chkCadastro(0).Value = 1 Then
                    rsSalvar.Fields(6) = 1
                End If
                
                If Check1.Value = 0 Then
                    rsSalvar.Fields(7) = "N"
                Else
                    rsSalvar.Fields(7) = "S"
                End If
                rsSalvar.Fields(8) = vCodcoligada 'Codigo da coligada
                rsSalvar.Fields(10) = txtCadastro(3) 'CodUsuario na TOTVS
                
                rsSalvar.Update

                SqlSalvar = "Select * from tbsenha where tbsenha.codigo = '" & Val(Me.mskCadastro(0)) & "'"
                rsSenha.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
                
                rsSenha.AddNew
                rsSenha.Fields(2) = mskCadastro(0).ClipText
                rsSenha.Fields(1) = txtCadastro(8)
                rsSenha.Fields(0) = txtCadastro(7).Text
                rsSenha.Fields(3) = vCodcoligada 'Codigo da coligada
                mobjMsg.Abrir "Inclusão realizada com sucesso!", Ok, informacao, "Atenção"
                LimpaControles
            Else
                rsSalvar.Fields(0) = mskCadastro(0).ClipText 'Código do usuário da IMRM
                rsSalvar.Fields(1) = txtCadastro(0).Text 'Nome do usuário da IMRM
                rsSalvar.Fields(4) = txtCadastro(6).Text 'Email do usuário da IMRM
                rsSalvar.Fields(2) = txtCadastro(1).Text 'codven na TOTVS
                rsSalvar.Fields(3) = txtCadastro(2).Text 'Nome do usuário na TOTVS
                rsSalvar.Fields(9) = Combo2.Text 'Codigo+nome Coligada TOTVS
                
                rsSalvar.Fields(5) = Val(Left(cboCadastro(1).Text, 6)) 'Grupo aoqual pertence o usuário da IMRM
                
                If chkCadastro(0).Value = 0 Then 'Informa se o usuário irá alterar a senha no próximo login
                    rsSalvar.Fields(6) = 0
                ElseIf chkCadastro(0).Value = 1 Then
                    rsSalvar.Fields(6) = 1
                End If
                
                If Check1.Value = 0 Then
                    rsSalvar.Fields(7) = "N"
                Else
                    rsSalvar.Fields(7) = "S"
                End If
                rsSalvar.Fields(8) = vCodcoligada 'Codigo da coligada
                rsSalvar.Fields(10) = txtCadastro(3) 'CodUsuario na TOTVS
                
                rsSalvar.Update
                
                SqlSalvar = "Select * from tbsenha where tbsenha.codigo = '" & Val(Me.mskCadastro(0)) & "'"
                rsSenha.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
                
                rsSenha.Fields(1) = txtCadastro(8).Text
                rsSenha.Fields(0) = txtCadastro(7).Text
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
    Exit Function
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Function
End Function

Private Sub gravaLocalEstoque()
    Dim rsSalvTrei As New ADODB.Recordset
    Dim SqlSalvTrei As String
    Dim Y As Integer, X As Integer
    
    SqlSalvTrei = "Delete from tblocalestoque where codigo= " & Val(Me.mskCadastro(0))
    rsSalvTrei.Open SqlSalvTrei, cnBanco
    
    SqlSalvTrei = "Select * from tblocalestoque as a where a.codigo = " & Val(Me.mskCadastro(0))
    rsSalvTrei.Open SqlSalvTrei, cnBanco, adOpenKeyset, adLockOptimistic
    
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        rsSalvTrei.AddNew
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsSalvTrei.Fields(0) = Val(Me.mskCadastro(0))
        rsSalvTrei.Fields(1) = ListView2.ListItems.Item(X)
        rsSalvTrei.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(1)
        rsSalvTrei.Fields(3) = vCodcoligada 'Codigo da coligada
    Next
    rsSalvTrei.Update
    rsSalvTrei.Close
    Set rsSalvTrei = Nothing
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
        ItemLst.SubItems(3) = txtCadastro(1).Text
        ItemLst.SubItems(4) = txtCadastro(2).Text
        ItemLst.SubItems(5) = txtCadastro(3).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(6) = ""
            ItemLst.ListSubItems.Item(6).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(6) = ""
            ItemLst.ListSubItems.Item(6).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadastro(0).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = Mid$(cboCadastro(1).Text, 10, 20)
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadastro(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtCadastro(2).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = txtCadastro(3).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
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
    
'    For X = 0 To txtCadastro.Count - 1
'        txtCadastro(X).Enabled = True
'    Next
'    For X = 0 To mskCadastro.Count - 1
'        mskCadastro(X).Enabled = True
'    Next
'    For X = 0 To cboCadastro.Count - 1
'        cboCadastro(X).Enabled = True
'    Next
'    chkCadastro(0).Enabled = True
    Combo1.Enabled = True
    mskCadastro(0).Enabled = False
End Function

Private Function BloqueiaControles()
    Dim X As Integer
'    For X = 0 To txtCadastro.Count - 1
'        txtCadastro(X).Enabled = False
'    Next
    For X = 0 To mskCadastro.Count - 1
        mskCadastro(X).Enabled = False
    Next
'    For X = 0 To cboCadastro.Count - 1
'        cboCadastro(X).Enabled = False
'    Next
'    chkCadastro(0).Enabled = False
    
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

Private Function chamaChapa(vNome As String)
On Error GoTo Err
    chamaChapa = False
    Dim rschamaChapa As New ADODB.Recordset
    Dim SqlchamaChapa As String
    
    If vNome = "" Then
'        SqlchamaChapa = "select a.CODVEN,a.NOME,a.CODCOLIGADA,F.NOMEFANTASIA,a.codusuario from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO INNER JOIN " & vBancoSAP & ".dbo.GCOLIGADA AS F ON A.CODCOLIGADA = F.CODCOLIGADA where a.CODCOLIGADA=1 and a.INATIVO=0 and a.codven = '" & txtCadastro(1).Text & "' AND C.CODSITUACAO in('A','F','P','Z') order by a.nome"
        SqlchamaChapa = "SELECT A.CODUSUARIO AS CODVEN,A.NOME,1 AS CODCOLIGADA,'IDG ENGENHARIA E CONSULTORIA LTDA' AS NOMEFANTASIA,A.CODUSUARIO FROM " & vBancoSAP & ".dbo.GUSUARIO AS A WHERE STATUS = 1"
    Else
'        SqlchamaChapa = "select a.CODVEN,a.NOME,a.CODCOLIGADA,F.NOMEFANTASIA,a.codusuario from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO INNER JOIN " & vBancoSAP & ".dbo.GCOLIGADA AS F ON A.CODCOLIGADA = F.CODCOLIGADA where a.CODCOLIGADA=1 and a.INATIVO=0 and a.nome like '" & vNome & "%' AND C.CODSITUACAO in('A','F','P','Z') order by a.nome"
        SqlchamaChapa = "SELECT A.CODUSUARIO AS CODVEN,A.NOME,1 AS CODCOLIGADA,'IDG ENGENHARIA E CONSULTORIA LTDA' AS NOMEFANTASIA,A.CODUSUARIO FROM " & vBancoSAP & ".dbo.GUSUARIO AS A WHERE A.STATUS = 1 AND A.NOME like '" & vNome & "%' ORDER BY A.NOME"
        ChamaGridChapa (SqlchamaChapa)
'        SqlchamaChapa = "select a.CODVEN,a.NOME,a.CODCOLIGADA,F.NOMEFANTASIA,a.codusuario from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO INNER JOIN " & vBancoSAP & ".dbo.GCOLIGADA AS F ON A.CODCOLIGADA = F.CODCOLIGADA where a.CODCOLIGADA=1 and a.INATIVO=0 and a.codven =" & Pesquisa & " AND C.CODSITUACAO in('A','F','P','Z') order by a.nome"
        SqlchamaChapa = "SELECT A.CODUSUARIO AS CODVEN,A.NOME,1 AS CODCOLIGADA,'IDG ENGENHARIA E CONSULTORIA LTDA' AS NOMEFANTASIA,A.CODUSUARIO FROM GUSUARIO AS A WHERE A.STATUS = 1 AND A.CODUSUARIO = " & Pesquisa & " ORDER BY A.NOME"
        vNome = ""
        'Exit Function
    End If
    rschamaChapa.Open SqlchamaChapa, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rschamaChapa.EOF Then
        txtCadastro(1).Text = Format(txtCadastro(1).Text, "000000")
        txtCadastro(2).Text = rschamaChapa.Fields(1)  'Nome
        txtCadastro(3).Text = rschamaChapa.Fields(4)  'Nome
'        CompoeControles = True
    Else
        mobjMsg.Abrir "Registro de colaborador não identificado no sistema", Ok, critico, "Atenção"
        txtCadastro(1).Text = ""
        txtCadastro(2).Text = "-"
        txtCadastro(3).Text = "-"
        txtCadastro(1).SetFocus
    End If
    rschamaChapa.Close
    Set rschamaChapa = Nothing
Err:
    Exit Function
End Function

Private Sub ChamaGridChapa(vSqlp As String)
    Dim F As New frmPesqger2
    If vSqlp = "" Then
        Sqlp = "select a.CODVEN,a.NOME,a.CODCOLIGADA,D.NOMEFANTASIA,a.codusuario from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b  on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA INNER JOIN " & vBancoSAP & ".dbo.GCOLIGADA AS D ON A.CODCOLIGADA = D.CODCOLIGADA where a.CODCOLIGADA=1 and a.INATIVO=0 AND C.CODSITUACAO in('A','F','P','Z')"
    Else
        Sqlp = vSqlp
        vSqlp = ""
    End If
    procnom = "nome"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Colaboradores"
    Pesquisa = frmUsuarios.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "codven=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If Pesquisa = "Pesquisa de Colaboradores" Then Pesquisa = ""
            txtCadastro(1) = Format(Pesquisa, "000000")
            txtCadastro(2) = rsLocal.Fields(1)
            txtCadastro(3) = rsLocal.Fields(4)
            Combo2.Text = Format(rsLocal.Fields(2), "000000") & " - " & rsLocal.Fields(3)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub ListView1_Click()
    MarcaDesmarcaGeral ListView1
End Sub

Private Sub ListView1_DblClick()
    addRemLoteNota ListView1, ListView2
End Sub

Private Sub ListView2_Click()
    MarcaDesmarcaGeral ListView2
End Sub

Private Sub ListView2_DblClick()
    addRemLoteNota ListView2, ListView1
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtCadastro(1).Text = Format(txtCadastro(1).Text, "000000")
            If chamaChapa("") = False Then Exit Sub
        End If
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            chamaChapa txtCadastro(2).Text
        End If
    End Select
End Sub
