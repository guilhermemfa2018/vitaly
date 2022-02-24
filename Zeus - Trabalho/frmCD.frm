VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Desenhos"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "frmCD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Status"
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
      Height          =   615
      Left            =   9600
      TabIndex        =   37
      Top             =   3840
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Status "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   35
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cboContDes 
         Height          =   315
         Index           =   0
         ItemData        =   "frmCD.frx":0CCA
         Left            =   120
         List            =   "frmCD.frx":0CD4
         TabIndex        =   1
         Tag             =   "Status"
         Text            =   "Aguardando"
         ToolTipText     =   "Status"
         Top             =   240
         Width           =   3015
      End
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   18
      Top             =   3840
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
      MICON           =   "frmCD.frx":0CEF
      PICN            =   "frmCD.frx":0D0B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3840
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
      MICON           =   "frmCD.frx":19E5
      PICN            =   "frmCD.frx":1A01
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
      Caption         =   "Registro de andamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5520
      TabIndex        =   30
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cboContDes 
         Height          =   315
         Index           =   2
         Left            =   2640
         TabIndex        =   12
         Tag             =   "Detalhista"
         ToolTipText     =   "Detalhista"
         Top             =   480
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "frmCD.frx":26DB
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtContDes 
         Height          =   1815
         Index           =   10
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   1680
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmCD.frx":274F
         TabIndex        =   36
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtContDes 
         Height          =   285
         Index           =   9
         Left            =   3480
         TabIndex        =   15
         Tag             =   "Croqui"
         ToolTipText     =   "Croqui"
         Top             =   1080
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Tag             =   "Data fim"
         ToolTipText     =   "Data fim"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   288555009
         CurrentDate     =   41366
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Tag             =   "Data início"
         ToolTipText     =   "Data início"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   288555009
         CurrentDate     =   41366
      End
      Begin VB.TextBox txtContDes 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Tag             =   "Usuário"
         ToolTipText     =   "Usuário"
         Top             =   480
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "frmCD.frx":27C3
         TabIndex        =   34
         Top             =   840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmCD.frx":2835
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmCD.frx":28A5
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmCD.frx":291B
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registro de entrada (Desenhos recebidos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtContDes 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3960
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "frmCD.frx":2989
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtContDes 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1080
         TabIndex        =   41
         Top             =   1080
         Width           =   4095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmCD.frx":29FD
         TabIndex        =   40
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         Top             =   480
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         ToolTipText     =   "Recebido"
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   162267137
         CurrentDate     =   41366
      End
      Begin VB.TextBox txtContDes 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   7
         Tag             =   "Peso Total"
         ToolTipText     =   "Peso Total"
         Top             =   1680
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "frmCD.frx":2A6B
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtContDes 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   6
         Tag             =   "Peso Unitário"
         ToolTipText     =   "Peso Unitário"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtContDes 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Tag             =   "Quantidade"
         ToolTipText     =   "Quantidade"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtContDes 
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   4
         Tag             =   "Revisão"
         ToolTipText     =   "Revisão"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtContDes 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "Desenho"
         ToolTipText     =   "Desenho"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtContDes 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "FCE"
         ToolTipText     =   "FCE"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Previsão de detalhamento "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   5055
         Begin VB.ComboBox cboContDes 
            Height          =   315
            Index           =   1
            ItemData        =   "frmCD.frx":2ADF
            Left            =   1440
            List            =   "frmCD.frx":2AE9
            TabIndex        =   10
            Tag             =   "Previsão de detalhamento"
            Text            =   "Dias"
            ToolTipText     =   "Previsão de detalhamento"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtContDes 
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   9
            Tag             =   "Pevisão de detalhamento"
            ToolTipText     =   "Pevisão de detalhamento"
            Top             =   360
            Width           =   1215
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "frmCD.frx":2AFA
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "frmCD.frx":2B6A
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmCD.frx":2BE8
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmCD.frx":2C54
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmCD.frx":2CC2
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmCD.frx":2D30
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificador "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtContDes 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Identificador"
         ToolTipText     =   "Identificador"
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsContDes As New ADODB.Recordset
Private sqlContDes As String
Private Status As String
Private rsLocal As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravarDados
'            gravaLog "Código esc.: " & txtContDes(0), "Nome esc: " & txtContDes(1), "Peso: " & txtContDes(2)
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela de cadastro?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Unload Me
            Set frmCD = Nothing
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

Private Sub Command1_Click()
    ChamaGridDesenho
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
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker3 = Date
    DTPicker1.CheckBox = True
    DTPicker2.CheckBox = True
    DTPicker3.CheckBox = True
    DTPicker1.Value = Null
    DTPicker2.Value = Null
    DTPicker3.Value = Null
    CompoeCombo cboContDes(2), "tbUsuarios", "codigo", "nome"
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        If cboContDes(0) = "Detalhado" Then
            BloqueiaControles
        Else
            DesbloqueiaControles
        End If
    End If
    'configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub GravarDados()
On Error GoTo Err
    If ValidaCampo = False Then Exit Sub
    Dim rsContDes As New ADODB.Recordset
    Dim sqlContDes As String
    Dim Y As Integer
10  cnBanco.BeginTrans
   
    sqlContDes = "select * from tbcd where idcd = '" & txtContDes(0) & "'"
    rsContDes.Open sqlContDes, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsContDes.EOF Then rsContDes.AddNew
    rsContDes.Fields(0) = Val(txtContDes(0))
    
    rsContDes.Fields(1) = "EXC" 'FCE
    rsContDes.Fields(2) = "EXC" 'Desenho
    rsContDes.Fields(3) = "EXC" 'Revisao
    rsContDes.Fields(17) = Val(txtContDes(12).Text) 'Identificador do desenho
    
    rsContDes.Fields(4) = txtContDes(4).Text
    rsContDes.Fields(5) = txtContDes(5).Text
    If Not IsNull(DTPicker1.Value) Then
        rsContDes.Fields(6) = DTPicker1.Value
    Else
        rsContDes.Fields(6) = Null
    End If
    rsContDes.Fields(7) = txtContDes(7).Text
    rsContDes.Fields(8) = cboContDes(1).Text
    rsContDes.Fields(9) = txtContDes(8).Text
    If Not IsNull(DTPicker2.Value) Then
        rsContDes.Fields(10) = DTPicker2.Value
    Else
        rsContDes.Fields(10) = Null
    End If
    If Not IsNull(DTPicker3.Value) Then
        rsContDes.Fields(11) = DTPicker3.Value
    Else
        rsContDes.Fields(11) = Null
    End If
    rsContDes.Fields(12) = txtContDes(9).Text
    If cboContDes(0).Text = "Aguardando" Then
        rsContDes.Fields(13) = 4
    Else
        rsContDes.Fields(13) = 5
    End If
    rsContDes.Fields(14) = txtContDes(10).Text
    If Check1.Value = 0 Then
        rsContDes.Fields(15) = "N"
    Else
        rsContDes.Fields(15) = "S"
    End If
    rsContDes.Fields(16) = cboContDes(2)
    rsContDes.Update
    cnBanco.CommitTrans
    rsContDes.Close
    Set rsContDes = Nothing
    AtualizaListview
    mobjMsg.Abrir "Os dados foram salvos com sucesso", Ok, informacao, "ZEUS"
    Unload Me
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
    End If
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DesbloqueiaControles
    For X = 0 To txtContDes.Count - 1
        txtContDes(X) = ""
    Next
    For X = 1 To cboContDes.Count - 1
        cboContDes(X) = ""
    Next
    txtContDes(0) = Format(GeraCodigo, "000000")
    txtContDes(8).Text = NomUsu 'Detalhista
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtContDes(0).Text = Format(rsContDes.Fields(0), "000000") 'IDCD
    If rsContDes.Fields(13) = 4 Then
        cboContDes(0).Text = "Aguardando" 'Status
    Else
        cboContDes(0).Text = "Detalhado" 'Status
    End If
    
    'VALORES ANTIGOS
    'txtContDes(1).Text = rsContDes.Fields(1) 'FCE
    If Not IsNull(rsContDes.Fields(21)) Then txtContDes(2).Text = rsContDes.Fields(21) 'Desenho
    If Not IsNull(rsContDes.Fields(22)) Then txtContDes(3).Text = rsContDes.Fields(22) 'Revisão
    
    'NOVOS VALORES
    If Not IsNull(rsContDes.Fields(17)) Then
        txtContDes(12).Text = Format(rsContDes.Fields(17), "000000") 'Identificador do Desenho
        txtContDes(1).Text = rsContDes.Fields(19) 'FCE
        txtContDes(11).Text = rsContDes.Fields(20) 'Projeto
    End If
    
    txtContDes(4).Text = rsContDes.Fields(4) 'Quantidade
    txtContDes(5).Text = rsContDes.Fields(5) 'Peso Unitário
    txtContDes(6).Text = rsContDes.Fields(4) * rsContDes.Fields(5) 'Peso Total
    If Not IsNull(rsContDes.Fields(6)) Then DTPicker1.Value = rsContDes.Fields(6) 'Data recebido
    txtContDes(7).Text = rsContDes.Fields(7) 'Tempo
    cboContDes(1).Text = rsContDes.Fields(8) 'horas/dias
    txtContDes(8).Text = rsContDes.Fields(9) 'Usuário
    If Not IsNull(rsContDes.Fields(10)) Then DTPicker2.Value = rsContDes.Fields(10) 'Data inicio
    If Not IsNull(rsContDes.Fields(11)) Then DTPicker3.Value = rsContDes.Fields(11) 'Data fim
    If Not IsNull(rsContDes.Fields(12)) Then txtContDes(9).Text = rsContDes.Fields(12) 'Croqui
    txtContDes(10).Text = rsContDes.Fields(14) 'Observação
    If rsContDes.Fields(15) = "S" Then 'Ativo
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If Not IsNull(rsContDes.Fields(16)) Then cboContDes(2).Text = rsContDes.Fields(16) 'Detalhista
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    For X = 0 To 8
        If txtContDes(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtContDes(X).Tag, Ok, critico, "Atenção"
            Me.txtContDes(X).SetFocus
            Exit Function
        End If
    Next
    
    For X = 0 To 1
        If txtContDes(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtContDes(X).Tag, Ok, critico, "Atenção"
            Me.txtContDes(X).SetFocus
            Exit Function
        End If
    Next
    If txtContDes(12).Text = "" Then
        mobjMsg.Abrir "Desenho ou Revisão não informados", Ok, critico, "Atenção"
        Exit Function
    End If
    
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 0 To txtContDes.Count - 1
        txtContDes(X).Enabled = False
    Next
    For X = 0 To cboContDes.Count - 1
        cboContDes(X).Enabled = False
    Next
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    DTPicker3.Enabled = False
    txtContDes(2).Enabled = True
    txtContDes(3).Enabled = True
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtContDes.Count - 1
        txtContDes(X).Enabled = True
    Next
    For X = 0 To cboContDes.Count - 1
        cboContDes(X).Enabled = True
    Next
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    DTPicker3.Enabled = True
    txtContDes(0).Enabled = False
    txtContDes(6).Enabled = False
    txtContDes(8).Enabled = False
    txtContDes(1).Enabled = False
    txtContDes(11).Enabled = False
    txtContDes(12).Enabled = False
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirContDes
    SqlGera = "Select top 1 * from tbCD order by idcd Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsContDes.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtContDes(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharContDes
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

Private Sub AbrirContDes()
On Error GoTo Err
    sqlContDes = "Select * from tbCD Order by idcd"
    rsContDes.Open sqlContDes, cnBanco, adOpenKeyset, adLockOptimistic
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

Private Sub FecharContDes()
    rsContDes.Close
    Set rsContDes = Nothing
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    sqlContDes = "select a.*,b.codprojeto,c.fce,c.projeto,b.desenho,b.revisao from tbcd as a left join tbdesenhos as b on a.iddesenho = b.iddesenho left join tbprojetos as c on b.codprojeto = c.codprojeto Where a.idcd= '" & Val(varGlobal) & "' order by a.idcd"
    rsContDes.Open sqlContDes, cnBanco, adOpenKeyset, adLockReadOnly
    If rsContDes.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Identificador não encontrado", Ok, critico, "Atenção"
    End If
    rsContDes.Close
    Set rsContDes = Nothing
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

Private Sub AtualizaListview()
    'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtContDes(0), "000000")) 'Identificador
        ItemLst.SubItems(1) = txtContDes(1).Text 'FCE
        ItemLst.SubItems(2) = txtContDes(2).Text 'Desenho
        ItemLst.SubItems(3) = txtContDes(3).Text 'Revisão
        ItemLst.SubItems(4) = txtContDes(4).Text 'Quantidade
        ItemLst.SubItems(5) = txtContDes(5).Text 'Peso Unit
        ItemLst.SubItems(6) = txtContDes(6).Text 'Peso Total
        If Not IsNull(DTPicker1.Value) Then
            ItemLst.SubItems(7) = DTPicker1.Value 'Recebido
        Else
            ItemLst.SubItems(7) = "-" 'Recebido
        End If
        ItemLst.SubItems(8) = txtContDes(7).Text & " " & cboContDes(1) 'Previsão detalhamento
        ItemLst.SubItems(9) = txtContDes(8).Text 'Usuário
        If Not IsNull(DTPicker2.Value) Then
            ItemLst.SubItems(10) = DTPicker2.Value 'Data inicio
        Else
            ItemLst.SubItems(10) = "-" 'Data inicio
        End If
        If Not IsNull(DTPicker3.Value) Then
            ItemLst.SubItems(11) = DTPicker3.Value 'Data fim
        Else
            ItemLst.SubItems(11) = "-" 'Data fim
        End If
        If txtContDes(9).Text <> "" Then ItemLst.SubItems(12) = txtContDes(9).Text 'Croqui
        If cboContDes(0).Text = "Aguardando" Then
            ItemLst.SubItems(13) = ""
            ItemLst.ListSubItems.Item(13).ReportIcon = "AGUARDE-02"
        Else
            ItemLst.SubItems(13) = ""
            ItemLst.ListSubItems.Item(13).ReportIcon = "OK"
        End If
        ItemLst.SubItems(14) = txtContDes(10).Text 'Observação
        If Check1.Value = 0 Then 'Ativo
            ItemLst.SubItems(15) = ""
            ItemLst.ListSubItems.Item(15).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(15) = ""
            ItemLst.ListSubItems.Item(15).ReportIcon = "OK"
        End If
        If cboContDes(2).Text <> "" Then ItemLst.SubItems(16) = cboContDes(2).Text 'Detalhista
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtContDes(1).Text 'FCE
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtContDes(2).Text 'Desenho
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtContDes(3).Text 'Revisão
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtContDes(4).Text 'Quantidade
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = txtContDes(5).Text 'Peso Unit
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = txtContDes(6).Text 'Peso Total
        If Not IsNull(DTPicker1.Value) Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = DTPicker1.Value 'Recebido
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = "-" 'Recebido
        End If
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = txtContDes(7).Text & " " & cboContDes(1) 'Previsão detalhamento
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = txtContDes(8).Text 'Usuário
        
        If Not IsNull(DTPicker2.Value) Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = DTPicker2.Value 'Data inicio
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = "-" 'Data inicio
        End If
        If Not IsNull(DTPicker2.Value) Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(11) = DTPicker3.Value 'Data fim
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(11) = "-" 'Data fim
        End If
        
        If txtContDes(9).Text <> "" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(12) = txtContDes(9).Text  'Croqui
        If cboContDes(0).Text = "Aguardando" Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13).ReportIcon = "AGUARDE-02"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13).ReportIcon = "OK"
        End If
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(14) = txtContDes(10).Text 'Observação
        If Check1.Value = 0 Then 'Ativo
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15).ReportIcon = "OK"
        End If
        If cboContDes(2).Text <> "" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(16) = cboContDes(2).Text 'Detalhista
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub configControles()
    If vSal = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub CalculaTotal()
On Error GoTo Err
    Dim valor1 As Double, valor2 As Double, valor3 As Double
    If txtContDes(4).Text <> "" Then valor1 = txtContDes(4).Text
    If txtContDes(5).Text <> "" Then valor2 = txtContDes(5).Text
    valor3 = valor1 * valor2
    If txtContDes(4).Text <> "" Or txtContDes(5).Text <> "" Then
        txtContDes(6).Text = Format(valor3, "#,##0.00;(#,##0.00)")
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Os Campo Quantidade ou Peso unitário possuem caracteres que não permitem o cálculo entre os mesmos", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub txtContDes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            achaDesenho
        End If
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            achaRevisao
        End If
    Case 4
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End Select

End Sub

Private Sub achaDesenho()
On Error GoTo Err
    Dim rsDesenho As New ADODB.Recordset
    Dim SqlDesenho As String
    Dim X As Integer
    
    SqlDesenho = "Select * from tbDesenhos where desenho = '" & txtContDes(2) & "' order by desenho"
    rsDesenho.Open SqlDesenho, cnBanco, adOpenKeyset, adLockOptimistic
    If rsDesenho.EOF Then
        txtContDes(2).Text = txtContDes(2)
        mobjMsg.Abrir "Desenho não cadastrado", Ok, critico, "Atenção"
    Else
        txtContDes(2).Text = rsDesenho.Fields(3)
    End If
    rsDesenho.Close
    Set rsDesenhoFCE = Nothing
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

Private Sub achaRevisao()
On Error GoTo Err
    Dim rsRevisao As New ADODB.Recordset
    Dim SqlRevisao As String
    
    SqlRevisao = "Select * from tbDesenhos where desenho = '" & txtContDes(2) & "' and revisao = '" & txtContDes(3) & "' order by desenho"
    rsRevisao.Open SqlRevisao, cnBanco, adOpenKeyset, adLockOptimistic
    If rsRevisao.EOF Then
        txtContDes(3).Text = txtContDes(3)
        mobjMsg.Abrir "Desenho não cadastrado", Ok, critico, "Atenção"
    Else
        txtContDes(3).Text = rsRevisao.Fields(4)
        achaFCEProj rsRevisao.Fields(2), rsRevisao.Fields(0)
    End If
    rsRevisao.Close
    Set rsRevisao = Nothing
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

Private Sub achaFCEProj(vCodProj As Integer, vIDDesenho)
On Error GoTo Err
    Dim rsProjeto As New ADODB.Recordset
    Dim SqlProjeto As String
    SqlProjeto = "Select * from tbProjetos where codprojeto = '" & vCodProj & "'"
    rsProjeto.Open SqlProjeto, cnBanco, adOpenKeyset, adLockOptimistic
    txtContDes(1).Text = rsProjeto.Fields(1)
    txtContDes(11).Text = rsProjeto.Fields(2)
    txtContDes(12).Text = Format(vIDDesenho, "000000")
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

Private Sub ChamaGridDesenho()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select a.desenho,a.revisao,a.codprojeto,a.iddesenho,b.projeto,b.fce from tbDesenhos as a inner join tbProjetos as b on a.codprojeto = b.codprojeto order by b.fce,b.projeto,a.desenho,a.revisao"
    procnom = "desenho"
    campo = 0
    Campo1 = 1
    campo2 = 4
    campo3 = 5
    Campo4 = 3
    Load F
    F.Caption = "Pesquisa de Desenho"
    Pesquisa = frmCD.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "iddesenho=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtContDes(2).Text = rsLocal.Fields(0)
            txtContDes(3).Text = rsLocal.Fields(1)
            txtContDes(12).Text = rsLocal.Fields(3)
            txtContDes(1).Text = rsLocal.Fields(5)
            txtContDes(11).Text = rsLocal.Fields(4)
            'achaFCEProj rsLocal.Fields(2), rsLocal.Fields(3)
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub txtContDes_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 5 Or Index = 4 Then
        'aceitar somente números e "Back Space", "Enter", "virgula"
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 And KeyAscii <> 45 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtContDes_LostFocus(Index As Integer)
    Select Case Index
    Case 4
        CalculaTotal
    Case 5
        CalculaTotal
    End Select
End Sub
