VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTerceirizados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Terceiros"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTerceirizados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCadTerc 
      Height          =   330
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtCadTerc 
      Height          =   330
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtCadTerc 
      Height          =   330
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtCadTerc 
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   9015
   End
   Begin VB.TextBox txtCadTerc 
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdCadastro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmTerceirizados.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "Salvar Critério"
      ToolTipText     =   "Salvar Critério"
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmTerceirizados.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   6000
      Width           =   615
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   8400
      TabIndex        =   25
      Top             =   6000
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   9375
      Begin VB.Frame Frame2 
         Caption         =   "Horário de trabalho"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   4560
         Width           =   6615
         Begin VB.TextBox txtCadTerc 
            Height          =   330
            Index           =   11
            Left            =   4080
            TabIndex        =   18
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCadTerc 
            Height          =   330
            Index           =   10
            Left            =   2760
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCadTerc 
            Height          =   330
            Index           =   9
            Left            =   1440
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCadTerc 
            Height          =   330
            Index           =   8
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmTerceirizados.frx":265E
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "frmTerceirizados.frx":26C6
            TabIndex        =   41
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   1440
            OleObjectBlob   =   "frmTerceirizados.frx":272A
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmTerceirizados.frx":279A
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtCadTerc 
         Height          =   330
         Index           =   7
         Left            =   2160
         TabIndex        =   12
         Top             =   3360
         Width           =   4095
      End
      Begin VB.TextBox txtCadTerc 
         Height          =   330
         Index           =   6
         Left            =   2160
         TabIndex        =   9
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtCadTerc 
         Height          =   330
         Index           =   5
         Left            =   2160
         TabIndex        =   6
         Top             =   1920
         Width           =   4095
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   6360
         TabIndex        =   13
         Top             =   3360
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   10
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   7
         Top             =   1920
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   6000
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   289275905
         CurrentDate     =   42416
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   6000
         OleObjectBlob   =   "frmTerceirizados.frx":2804
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   4200
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   289275905
         CurrentDate     =   42416
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   289275905
         CurrentDate     =   42416
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "frmTerceirizados.frx":287E
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "frmTerceirizados.frx":28F2
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmTerceirizados.frx":296C
         Left            =   120
         List            =   "frmTerceirizados.frx":2979
         TabIndex        =   14
         Top             =   4080
         Width           =   6615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTerceirizados.frx":2999
         TabIndex        =   34
         Top             =   3840
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTerceirizados.frx":2A01
         TabIndex        =   33
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Foto "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Index           =   0
         Left            =   6840
         TabIndex        =   30
         Top             =   1560
         Width           =   2415
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   1
            Left            =   720
            Picture         =   "frmTerceirizados.frx":2A79
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3120
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   0
            Left            =   120
            Picture         =   "frmTerceirizados.frx":3743
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3120
            Width           =   615
         End
         Begin VB.PictureBox Picture2 
            Height          =   2655
            Left            =   120
            ScaleHeight     =   2595
            ScaleWidth      =   2115
            TabIndex        =   31
            Top             =   240
            Width           =   2175
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   2655
               Left            =   -120
               Top             =   0
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   4683
               Image           =   "frmTerceirizados.frx":440D
            End
            Begin VB.Label Label59 
               Alignment       =   2  'Center
               Caption         =   "A Imagem não se encontra no local especificado"
               Height          =   495
               Left            =   120
               TabIndex        =   32
               Top             =   1200
               Visible         =   0   'False
               Width           =   2055
            End
         End
         Begin MSComDlg.CommonDialog cdlFoto 
            Left            =   1800
            Top             =   3240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTerceirizados.frx":4425
         TabIndex        =   29
         Top             =   2400
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTerceirizados.frx":448B
         TabIndex        =   28
         Top             =   1680
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTerceirizados.frx":44EF
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTerceirizados.frx":4553
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label53 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   4920
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmTerceirizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTerceirizados As New ADODB.Recordset
Private SqlTerceirizados As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        'carregar imagem para o Picture
        With cdlFoto
            .Filter = "(Arquivo *.JPG)|*.jpg"
            .ShowOpen
            Caminho1 = .FileName
        End With
        'mostra a figura
        'Image1.Picture = LoadPicture(Caminho1)
        aicAlphaImage1.LoadImage_FromFile (Caminho1)
        Label53 = Caminho1
    
    Case 1
        aicAlphaImage1.ClearImage
        Label53 = ""
    Case 2
        ChamaGrid "CORPORERM.dbo.PSECAO", "descricao", txtCadTerc(2), frmTerceirizados, "codigo", "descricao"
        CarregaTxt "CORPORERM.dbo.PSECAO", "codigo", "S", "", "", txtCadTerc(2), txtCadTerc(5), 1, 2, txtCadTerc(2), "S", txtCadTerc(5), "1"
        txtCadTerc(2).SetFocus
    Case 3
        ChamaGrid "CORPORERM.dbo.PFUNCAO", "nome", txtCadTerc(3), frmTerceirizados, "codigo", "nome"
        CarregaTxt "CORPORERM.dbo.PFUNCAO", "codigo", "S", "", "", txtCadTerc(3), txtCadTerc(6), 1, 2, txtCadTerc(3), "S", txtCadTerc(6), "1"
        txtCadTerc(3).SetFocus
    Case 6
        ChamaGrid "CORPORERM.dbo.GCCUSTO", "nome", txtCadTerc(4), frmTerceirizados, "codigo", "nome"
        CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtCadTerc(4), txtCadTerc(7), 7, 2, txtCadTerc(4), "S", txtCadTerc(7), "1"
        txtCadTerc(4).SetFocus
    Case 4
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "SAF"
        Else
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "SAF"
        End If
    
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
    Status = Pesquisa
    DTPicker1 = Date
    DTPicker2 = Date
    If Status = "novo" Then
        txtCadTerc(0).Text = "CONTR" & Format(GeraCodigoTB("tbTerceirizados", "chapa", "TERC", ""), "000000")
    ElseIf Status = "editar" Then
        ResultPesq
    End If
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    Me.Top = (Principal.Height / 2) - (Me.Height / 2)
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtCadTerc_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtCadTerc(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtCadTerc.Count - 1
        txtCadTerc(X).SelStart = 0
        txtCadTerc(X).SelLength = Len(txtCadTerc(X).Text)
    Next
End Sub

Private Sub txtCadTerc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    If Index = 2 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaTxt "CORPORERM.dbo.PSECAO", "codigo", "S", "", "", txtCadTerc(2), txtCadTerc(5), 1, 2, txtCadTerc(2), "S", txtCadTerc(5), "1"
            'CarregaSecao
        End If
    End If
    CampoHora txtCadTerc(8), KeyCode
    CampoHora txtCadTerc(9), KeyCode
    CampoHora txtCadTerc(10), KeyCode
    CampoHora txtCadTerc(11), KeyCode
Error:
    Exit Sub
End Sub

'Private Sub txtCadTerc_KeyPress(index As Integer, KeyCode As Integer, Shift As Integer)
'   CampoHora txtCadTerc(8), KeyCode
'End Sub

Private Sub txtCadTerc_LostFocus(Index As Integer)
    voltaCorText txtCadTerc(Index)
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    SqlTerceirizados = "select a.chapa,a.nome,a.idsetor,a.setor,a.idfuncao,a.funcao,a.idcc,a.nmcc,a.empresa,a.datacadastro,a.datacontratoini,a.datacontratofim,a.ativo,a.foto,a.hentrada,a.rinicio,a.rfim,a.hsaida from tbTerceirizados as a where a.chapa = '" & varGlobal & "'"
    rsTerceirizados.Open SqlTerceirizados, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTerceirizados.RecordCount > 0 Then
        compoeControlesForm
    End If
    rsTerceirizados.Close
    Set rsTerceirizados = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Sub

Private Sub compoeControlesForm()
    txtCadTerc(0) = Format(rsTerceirizados.Fields(0), "000000") 'Chapa
    txtCadTerc(1) = rsTerceirizados.Fields(1) 'Nome
    txtCadTerc(2) = rsTerceirizados.Fields(2) 'ID do Setor
    txtCadTerc(5) = rsTerceirizados.Fields(3) 'Nome do Setor
    txtCadTerc(3) = rsTerceirizados.Fields(4) 'Id da Função
    txtCadTerc(6) = rsTerceirizados.Fields(5) 'Nome da Função
    txtCadTerc(4) = rsTerceirizados.Fields(6) 'ID do Centro de Custo
    txtCadTerc(7) = rsTerceirizados.Fields(7) 'Nome do Centro de Custo
    Combo1.Text = rsTerceirizados.Fields(8) 'nome da empresa que contratou o colaborador
    DTPicker1.Value = rsTerceirizados.Fields(9) 'Data de Cadastro do Colaborador
    DTPicker2.Value = rsTerceirizados.Fields(10) 'Data do inicio do contrato do colaborador
    If Not IsNull(rsTerceirizados.Fields(11)) Then
        DTPicker3.Value = rsTerceirizados.Fields(11) 'Data do Fim do contrato do colaborador
    End If
    
    txtCadTerc(8) = Mid$(rsTerceirizados.Fields(14), 1, 5) ' Horario de inicio do expediente
    txtCadTerc(9) = Mid$(rsTerceirizados.Fields(15), 1, 5) ' Horário de inicio do almoço
    txtCadTerc(10) = Mid$(rsTerceirizados.Fields(16), 1, 5) ' Horário de fim do almoço
    txtCadTerc(11) = Mid$(rsTerceirizados.Fields(17), 1, 5) ' Horário de fim do expediente
    
    
    Label53 = rsTerceirizados.Fields(13) 'Local onde esta armazenado a foto do coloborador
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    If rsTerceirizados.Fields(12) = "N" Or IsNull(rsTerceirizados.Fields(12)) Then
        Check1.Value = 0 'Status (não ativo)
    Else
        Check1.Value = 1 'Status (ativo)
    End If
End Sub

Private Function salvar_Dados()
'On Error GoTo Err
    salvar_Dados = True
    Dim X As Integer, Y As Integer
    
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    limpaQualquerDado
    
    
    vQualquerDado(1, 1) = txtCadTerc(0).Text 'Chapa terceirizado
    vQualquerDado(1, 2) = "S"
    vQualquerDado(2, 1) = txtCadTerc(1).Text 'Nome do terceirizado
    vQualquerDado(2, 2) = "S"
    vQualquerDado(3, 1) = Label53.Caption ' Descrição do critério
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtCadTerc(2).Text 'id do setor
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtCadTerc(5).Text 'nome do setor
    vQualquerDado(5, 2) = "S"
    vQualquerDado(6, 1) = txtCadTerc(3).Text 'id da função
    vQualquerDado(6, 2) = "S"
    vQualquerDado(7, 1) = txtCadTerc(6).Text 'nome do setor
    vQualquerDado(7, 2) = "S"
    vQualquerDado(8, 1) = txtCadTerc(4).Text 'id do centro de custo
    vQualquerDado(8, 2) = "S"
    vQualquerDado(9, 1) = txtCadTerc(7).Text 'nome do centro de custo
    vQualquerDado(9, 2) = "S"
    vQualquerDado(10, 1) = Combo1.Text 'nome da empresa que contratou o colaborador
    vQualquerDado(10, 2) = "S"
    If Check1.Value = 1 Then
        vQualquerDado(11, 1) = "S" ' Status do critério
    Else
        vQualquerDado(11, 1) = "N" ' Status do critério
    End If
    vQualquerDado(11, 2) = "S"
    vQualquerDado(12, 1) = DTPicker1.Value 'Data de cadastro do colaborador
    vQualquerDado(12, 2) = "D"
    vQualquerDado(13, 1) = DTPicker2.Value 'Data do inicio do contrato do colaborador
    vQualquerDado(13, 2) = "D"
    If DTPicker3.Value <> "" Then
        vQualquerDado(14, 1) = DTPicker3.Value 'Data do fim do contrato do colaborador
    'Else
    '    vQualquerDado(14, 1) = Null 'Data do fim do contrato do colaborador
    End If
    vQualquerDado(14, 2) = "D"
    
    vQualquerDado(15, 1) = txtCadTerc(8) 'Horário de inicio do expediente
    vQualquerDado(15, 2) = "S"
    vQualquerDado(16, 1) = txtCadTerc(9) 'Horário de inicio refeição
    vQualquerDado(16, 2) = "S"
    vQualquerDado(17, 1) = txtCadTerc(10) 'Horário de fim refeição
    vQualquerDado(17, 2) = "S"
    vQualquerDado(18, 1) = txtCadTerc(11) 'Horário de fim do expediente
    vQualquerDado(18, 2) = "S"
    
    GravaDados "tbTerceirizados", "chapa", "S", txtCadTerc(0), 18, "", "", txtCadTerc(0)
        
    'AtualizaListview
    Exit Function
Err:
    salvar_Dados = False
End Function

