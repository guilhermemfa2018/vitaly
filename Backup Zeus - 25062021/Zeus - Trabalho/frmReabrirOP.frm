VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Begin VB.Form frmReabrirOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reabrir Operação"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReabrirOP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   12015
   Begin VB.Frame Frame7 
      Caption         =   "Justificativa para reabertura "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   33
      Top             =   4440
      Width           =   6855
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmReabrirOP.frx":0CCA
         Left            =   120
         List            =   "frmReabrirOP.frx":0CD1
         TabIndex        =   10
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dados da baixa:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   7080
      TabIndex        =   24
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   12
         Left            =   2760
         TabIndex        =   15
         Tag             =   "Baixas indevidas realizadas em qualquer Centro de Custo"
         ToolTipText     =   "Baixas indevidas realizadas em qualquer Centro de Custo"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   11
         Left            =   2760
         TabIndex        =   14
         Tag             =   "Baixas indevidas realizadas no PRÓPRIO Centro de Custo"
         ToolTipText     =   "Baixas indevidas realizadas no PRÓPRIO Centro de Custo"
         Top             =   2280
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmReabrirOP.frx":0CF1
         TabIndex        =   39
         Top             =   2640
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmReabrirOP.frx":0D77
         TabIndex        =   38
         Top             =   2040
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":0DF7
         TabIndex        =   37
         Top             =   3720
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":0E65
         TabIndex        =   32
         Top             =   3360
         Width           =   4575
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   10
         Left            =   2760
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   2760
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         Caption         =   "Colaborador "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2535
         Begin VB.PictureBox Picture1 
            Height          =   2655
            Left            =   120
            ScaleHeight     =   2595
            ScaleWidth      =   2235
            TabIndex        =   26
            Top             =   240
            Width           =   2295
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   2655
               Left            =   0
               Top             =   0
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   4683
               Image           =   "frmReabrirOP.frx":0EE3
            End
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmReabrirOP.frx":0EFB
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   2760
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   106430465
         CurrentDate     =   41967
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmReabrirOP.frx":0F61
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmReabrirOP.frx":0FC3
         TabIndex        =   31
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tempos "
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
      Left            =   3600
      TabIndex        =   23
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmReabrirOP.frx":1025
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":1093
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status da Operação: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      TabIndex        =   22
      Top             =   4440
      Width           =   4815
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":10F9
         TabIndex        =   34
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Operação "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   6855
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   375
         Index           =   13
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmReabrirOP.frx":1153
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   5160
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   1095
         Index           =   8
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   6615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":11BB
         TabIndex        =   28
         Top             =   1680
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frmReabrirOP.frx":1249
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   6615
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtReabrirOP 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":12BB
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "frmReabrirOP.frx":132F
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReabrirOP.frx":139F
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frmReabrirOP.frx":1403
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtReabrirOP 
         Height          =   330
         Index           =   14
         Left            =   2280
         TabIndex        =   43
         Top             =   960
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Código de barras da operação "
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
      TabIndex        =   16
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtReabrirOP 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
   End
   Begin ZEUS.chameleonButton cmdCD 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   35
      Top             =   5400
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
      MICON           =   "frmReabrirOP.frx":1479
      PICN            =   "frmReabrirOP.frx":1495
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdCD 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Tag             =   "Reabrir Operação"
      ToolTipText     =   "Reabrir Operação"
      Top             =   5400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmReabrirOP.frx":216F
      PICN            =   "frmReabrirOP.frx":218B
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
Attribute VB_Name = "frmReabrirOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vIDCD As TextBox
Private vObservacao As TextBox
Private vEmailAprovador As String

Private Sub cmdCD_Click(Index As Integer)
    Select Case Index
    Case 0
        If Combo1.Text = "" Then
            mobjMsg.Abrir "Selecione uma JUSTIFICATIVA para Reabrir a operação", Ok, critico, "Atenção"
            Combo1.SetFocus
            Exit Sub
        End If
        mobjMsg.Abrir "Confirma a reabertura dessa operação", YesNo, pergunta, "Zeus"
        If Tp = 1 Then
            'Rotina em Desenvolvimento
            If ReabrirOperacao(Val(txtReabrirOP(4).Text), Val(txtReabrirOP(3).Text), Val(txtReabrirOP(13).Text), Val(txtReabrirOP(5).Text)) = True Then
                mobjMsg.Abrir "Operação REABERTA com sucesso!", Ok, informacao, "ZEUS"
                If dadosEmail = False Then Exit Sub
                If vSMTP <> "" Then enviaEmail
                Unload Me
            Else
                mobjMsg.Abrir "Erro ao REABRIR operação", Ok, critico, "ZEUS"
            End If
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Set vIDCD = Me.Controls.Add("VB.TextBox", "vIDCD")
    Set vObservacao = Me.Controls.Add("VB.TextBox", "vObservacao")
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    Me.Top = 800
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub txtReabrirOP_GotFocus(Index As Integer)
    mudaCorText txtReabrirOP(Index)
End Sub

Private Sub txtReabrirOP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        If txtReabrirOP(0).Text <> "" Then
            CompoeDados txtReabrirOP(0).Text
        End If
    End If
End Sub

Private Sub txtReabrirOP_LostFocus(Index As Integer)
    voltaCorText txtReabrirOP(Index)
    CompoeDados txtReabrirOP(0).Text
End Sub

Private Sub CompoeDados(vCBarra As String)
    On Error Resume Next
    If txtReabrirOP(0).Text = "" Then Exit Sub
    Dim rsAchaOS_Prog As New ADODB.Recordset
    Dim SqlAchaOS_Prog As String
    
    Dim rsCompoeDados As New ADODB.Recordset
    Dim SqlCompoeDados As String
    Dim vOS As String, vProgramacao As String, vLocaliza As String
    Dim vChapa As String
    
    txtReabrirOP(5).Text = ""
    SkinLabel11.Caption = ""
    calculaTempoApropriado txtReabrirOP(0).Text
    
    SqlAchaOS_Prog = "select a.idprogramacao,a.idos,a.idoperacao,a.status from tbMPItens as a where a.codigobarra = '" & vCBarra & "'"
    rsAchaOS_Prog.Open SqlAchaOS_Prog, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaOS_Prog.RecordCount = 0 Then
        mobjMsg.Abrir "Código de Barras não encontrado", Ok, critico, "Atenção"
        rsAchaOS_Prog.Close
        Set rsAchaOS_Prog = Nothing
        Exit Sub
    Else
        vProgramacao = rsAchaOS_Prog.Fields(0)
        vOS = rsAchaOS_Prog.Fields(1)
        vLocaliza = vProgramacao & vOS
        txtReabrirOP(5).Text = rsAchaOS_Prog.Fields(2)
        If rsAchaOS_Prog.Fields(3) = 3 Then
            SkinLabel11.Caption = "BAIXADA"
            cmdCD(0).Enabled = True
        Else
            SkinLabel11.Caption = "ABERTA"
            cmdCD(0).Enabled = False
        End If
    End If
    rsAchaOS_Prog.Close
    Set rsAchaOS_Prog = Nothing
    
    'ABAIXO: Verifica se existe baixa para a operação digitada no código de barras
    'SqlCompoeDados = "select dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado,c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,c.observacao,CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai," & _
    '                 "b.chapa,b.NOME,f.NOME as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado from CORPORERM.dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA left join CORPORERM.dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join CORPORERM.dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO " & _
    '                 "left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vCBarra & "' and a.idparada = '9020'"
    SqlCompoeDados = "Select dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado,c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,substring(c.observacao,1,300) as observacao,CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai,b.chapa,b.NOME,f.NOME as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado " & _
                     "from " & vBancoTotvs & ".dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vCBarra & "' and a.idparada = '9020' union " & _
                     "Select dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado,c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,substring(c.observacao,1,300) as observacao,CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai,b.chapa COLLATE SQL_Latin1_General_CP1_CI_AI CHAPA,b.NOME COLLATE SQL_Latin1_General_CP1_CI_AI as NOME, " & _
                     "b.nmcc COLLATE SQL_Latin1_General_CP1_CI_AI as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado from tbTerceirizados as b left join tbOsMov as a on a.chapa = b.CHAPA left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vCBarra & "' and a.idparada = '9020'"
    
    
    rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockReadOnly
    'ABAIXO: Se não houver baixa o sistema procura em qual operação houve a baixa
    If rsCompoeDados.RecordCount = 0 Then
        rsCompoeDados.Close
        Set rsCompoeDados = Nothing
        
        'SqlCompoeDados = "select dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado,c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,c.observacao,CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai," & _
        '                 "b.chapa,b.NOME,f.NOME as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado from " & vBancoTotvs & ".dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO " & _
        '                 "left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vCBarra & "'"
        
        SqlCompoeDados = "select dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado,c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,substring(c.observacao,1,300),CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai,b.chapa,b.NOME,f.NOME as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado " & _
                         "from " & vBancoTotvs & ".dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vCBarra & "' union select dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado, " & _
                         "c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,substring(c.observacao,1,300),CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai,b.chapa COLLATE SQL_Latin1_General_CP1_CI_AI chapa,b.NOME COLLATE SQL_Latin1_General_CP1_CI_AI nome,b.nmcc COLLATE SQL_Latin1_General_CP1_CI_AI as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado " & _
                         "from tbTerceirizados as b left join tbOsMov as a on a.chapa = b.CHAPA left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vCBarra & "'"
        
        
        rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockReadOnly
        If rsCompoeDados.RecordCount > 0 Then txtReabrirOP(5).Text = rsCompoeDados.Fields(3)
        
        rsCompoeDados.Close
        Set rsCompoeDados = Nothing
        
        SqlCompoeDados = "select top 1 dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Orçado,c.idos,c.idprogramacao,c.idoperacao,CONVERT (VARCHAR, c.dataprevista, 103) as dataProgramada,c.grupo,c.observacao,CONVERT (VARCHAR, c.databaixa, 103) as dataBaixa,CONVERT (VARCHAR, a.horasai, 108) as horasai," & _
                         "b.chapa,b.NOME,f.NOME as sub_centro,a.codigobarra,c.status,c.revisaoos,c.nomecc as ccbaixado from " & vBancoTotvs & ".dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO " & _
                         "left join tbMPItens as c on a.codigobarra = c.codigobarra where a.codigobarra like '" & vLocaliza & "%'" & " and a.idparada = '9020' order by c.idos,c.idprogramacao,c.idoperacao desc"
        rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    
    If rsCompoeDados.RecordCount > 0 Then
        txtReabrirOP(1).Text = rsCompoeDados.Fields(0) 'Tempo Orçado
        'txtReabrirOP(2).Text = "" 'Tempo Apropriado
        txtReabrirOP(3).Text = rsCompoeDados.Fields(1) 'OS
        txtReabrirOP(13).Text = rsCompoeDados.Fields(14) 'Nº de Revisão da OS
        txtReabrirOP(4).Text = rsCompoeDados.Fields(2) 'Programação
        If txtReabrirOP(5).Text = "" Then txtReabrirOP(5).Text = rsCompoeDados.Fields(3) 'Operação
        txtReabrirOP(6).Text = DatePart("ww", rsCompoeDados.Fields(4), vbMonday, vbFirstFourDays) 'Semana Programacao
        txtReabrirOP(7).Text = rsCompoeDados.Fields(5) 'Nome da Operação
        txtReabrirOP(14).Text = rsCompoeDados.Fields(15) ' Centro de Custo Baixado
        txtReabrirOP(8).Text = rsCompoeDados.Fields(6) 'Observação do Planejamento
        txtReabrirOP(9).Text = DatePart("ww", rsCompoeDados.Fields(7), vbMonday, vbFirstFourDays) 'Semana da Baixa
        DTPicker1.Value = rsCompoeDados.Fields(7) 'Data da Baixa
        txtReabrirOP(10).Text = rsCompoeDados.Fields(8) 'Hora da Baixa
        SkinLabel10.Caption = rsCompoeDados.Fields(9) & " - " & rsCompoeDados.Fields(10) 'Chapa/Nome Colaborador que efetuou a baixa indevida
        SkinLabel12.Caption = rsCompoeDados.Fields(11) 'Subcentro de trabalho do colaborador
        
        If SkinLabel11.Caption = "" Then
            If rsCompoeDados.Fields(13) = 3 Then
                SkinLabel11 = "BAIXADA"
                cmdCD(0).Enabled = True
            Else
                SkinLabel11 = "ABERTA"
                cmdCD(0).Enabled = False
            End If
        End If
        vChapa = rsCompoeDados.Fields(9)
    Else
        
        'Verificar se encontra-se baixada diretamente pela Inspeção
        If SkinLabel11 <> "ABERTA" Then
            Dim rsVerBaixaInsp As New ADODB.Recordset
            Dim SqlVerBaixaInsp As String
            SqlVerBaixaInsp = "select * from tbositens where codigobarra ='" & vCBarra & "'"
            rsVerBaixaInsp.Open SqlVerBaixaInsp, cnBanco, adOpenKeyset, adLockReadOnly
            If rsVerBaixaInsp.RecordCount = 0 Then
                mobjMsg.Abrir "Dados NÃO encontrados ou OPERAÇÃO já foi reaberta", Ok, critico, "Atenção"
                rsCompoeDados.Close
                Set rsCompoeDados = Nothing
                Exit Sub
            Else
                txtReabrirOP(3).Text = rsVerBaixaInsp.Fields(0) 'OS
                txtReabrirOP(13).Text = rsVerBaixaInsp.Fields(1) 'Nº de Revisão da OS
                txtReabrirOP(4).Text = rsVerBaixaInsp.Fields(7) 'Programação
                If txtReabrirOP(5).Text = "" Then txtReabrirOP(5).Text = rsVerBaixaInsp.Fields(10) 'Operação
                txtReabrirOP(14).Text = rsVerBaixaInsp.Fields(6) ' Centro de Custo Baixado
                txtReabrirOP(8).Text = "OBSERVAÇÃO DO SISTEMA: Essa operação foi baixada pelo Controle de Qualidade"
            End If
        Else
            mobjMsg.Abrir "OPERAÇÃO já encontra-se aberta", Ok, critico, "Atenção"
            txtReabrirOP(8).Text = ""
        End If
    End If
    rsCompoeDados.Close
    Set rsCompoeDados = Nothing
    
    'Encontra quantidade de baixas indevidas no centro de custo do colaborados
    SqlCompoeDados = "select * from tbbaixasindevidasOP where chapa ='" & Mid$(SkinLabel10, 1, 5) & "' and centrodecustobaixa ='" & txtReabrirOP(12).Text & "'"
    rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCompoeDados.RecordCount > 0 Then txtReabrirOP(11).Text = rsCompoeDados.RecordCount Else txtReabrirOP(11).Text = "0"
    rsCompoeDados.Close
    Set rsCompoeDados = Nothing
    
    'Encontra quantidade de baixas indevidas em qualquer centro de custo
    SqlCompoeDados = "select * from tbbaixasindevidasOP where chapa ='" & Mid$(SkinLabel10, 1, 5) & "'"
    rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCompoeDados.RecordCount > 0 Then txtReabrirOP(12).Text = rsCompoeDados.RecordCount Else txtReabrirOP(12).Text = "0"
    rsCompoeDados.Close
    Set rsCompoeDados = Nothing
    
'------------------------
    SqlCompoeDados = "select c.IDIMAGEM,a.chapa,a.nome,b.IMAGEM from " & vBancoTotvs & ".dbo.PFUNC as a left join " & vBancoTotvs & ".dbo.PPESSOA as c on a.CODPESSOA = c.CODIGO left join " & vBancoTotvs & ".dbo.GIMAGEM as b on c.IDIMAGEM = b.ID " & _
                "where a.CHAPA = '" & vChapa & "'  order by a.nome"
    rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockReadOnly
    Dim mStream As ADODB.Stream
    Set mStream = New ADODB.Stream
    mStream.Type = adTypeBinary
    mStream.Open
    mStream.Write rsCompoeDados.Fields(3).Value
    mStream.SaveToFile App.Path & "\Temp1.jpg", adSaveCreateOverWrite
    aicAlphaImage1.ClearImage
    aicAlphaImage1.LoadImage_FromFile (App.Path & "\temp1.jpg")
    Kill App.Path & "\Temp1.jpg"
    
    rsCompoeDados.Close
    Set rsCompoeDados = Nothing

End Sub


Private Sub calculaTempoApropriado(vCBarra As String)
    Dim rsHAprop As New ADODB.Recordset
    Dim sqlHAprop As String
    Dim vHorasApropriadas As String
    
    sqlHAprop = "select CONVERT (VARCHAR, a.horasai-a.horaent, 108) as horaent from tbOsMov  as a where a.codigobarra = '" & vCBarra & "'"
    rsHAprop.Open sqlHAprop, cnBanco, adOpenKeyset, adLockReadOnly
    vHorasApropriadas = "0000:00"
    Do While Not rsHAprop.EOF
        If Not IsNull(rsHAprop.Fields(0)) Then somaTempoPPSAtraso rsHAprop.Fields(0), vHorasApropriadas
        rsHAprop.MoveNext
    Loop
    rsHAprop.Close
    Set rsHAprop = Nothing
    txtReabrirOP(2).Text = vHorasApropriadas
End Sub

Private Function somaTempoPPSAtraso(vTempo, vOndeAcumula As String)
    Dim seg As Long, min As Long, hora As Long
    Dim tempo As Long
    Dim matriz2

    matriz2 = Split(vTempo, ":")
    tempo = tempo + (CLng(matriz2(0)) * 3600)
    tempo = tempo + (CLng(matriz2(1)) * 60)
    
    If vOndeAcumula <> "" Then
        matriz2 = Split(vOndeAcumula, ":")
        tempo = tempo + (CLng(matriz2(0)) * 3600)
        tempo = tempo + (CLng(matriz2(1)) * 60)
    End If
    
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    
    vOndeAcumula = Format(hora, "0000") & ":" & Format(min, "00")
    somaTempoPPSAtraso = vOndeAcumula
End Function


Private Function ReabrirOperacao(vProgramacao As Integer, vOS As Integer, vRev As Integer, vOP As Integer)
'On Error GoTo Err
    ReabrirOperacao = True
    Dim rsReabreOP As New ADODB.Recordset
    Dim SqlReabreOP As String
    
    Dim rsRegBaixaInd As New ADODB.Recordset
    Dim SqlRegBaixaInd As String
    
    'cnBanco.BeginTrans
    
    vTransacaoAtiva = cnBanco.BeginTrans
    
    'ABAIXO: Realiza reabertura da operação no sistema
    SqlReabreOP = "update tbMPItens set status = 2 where idprogramacao = '" & vProgramacao & "' and idos = '" & vOS & "' and idoperacao = '" & vOP & "' and revisaoos = '" & vRev & "'"
    rsReabreOP.Open SqlReabreOP, cnBanco
    
    SqlReabreOP = "update tbos set status = 2 where idos = '" & vOS & "' and revisao = '" & vRev & "'"
    rsReabreOP.Open SqlReabreOP, cnBanco
    
    SqlReabreOP = "update tbositens set status = 2 where idos = '" & vOS & "' and idoperacao = '" & vOP & "'"
    rsReabreOP.Open SqlReabreOP, cnBanco
    
    SqlReabreOP = "Update tbMP set status = 2 where idprogramacao = '" & vProgramacao & "'"
    rsReabreOP.Open SqlReabreOP, cnBanco
    
    'ABAIXO: Registra ocorrência de baixa indevida para o colaborador
    SqlRegBaixaInd = "Insert into tbBaixasIndevidasOP(chapa,nome,codigobarra,centrodecustocolab,centrodecustobaixa,reabertapor,datareabertura,horareabertura) " & _
                     "values('" & Mid$(SkinLabel10, 1, 5) & "','" & Mid$(SkinLabel10, 9, 50) & "','" & txtReabrirOP(0).Text & "','" & SkinLabel12 & "','" & txtReabrirOP(14).Text & "','" & NomUsu & "','" & Format(CDate(vDataDoBanco), "YYYY-MM-DD") & "','" & Time & "')"
    rsRegBaixaInd.Open SqlRegBaixaInd, cnBanco
    
    'ABAIXO: Altera o código de parada da baixa realizada para ERRO
    SqlReabreOP = "Update tbOsMov set idparada = 'ERRO' where codigobarra = '" & txtReabrirOP(0).Text & "' and idparada = '9020'"
    rsReabreOP.Open SqlReabreOP, cnBanco
    
    GravaCD
    If GravaCD = False Then
        GoTo Err
    End If
    
    If vTransacaoAtiva > 0 Then cnBanco.CommitTrans
    Exit Function
Err:
    cnBanco.RollbackTrans
    ReabrirOperacao = False
End Function


Private Function GravaCD()
'On Error GoTo Err
    GravaCD = True
    
    vObservacao = "Foi realizada a baixa indevida da OP mencionada abaixo:" & vbCrLf & _
    " OS nº: " & txtReabrirOP(3) & "/" & txtReabrirOP(13) & " " & vbCrLf & _
    " Programação nº: " & txtReabrirOP(4) & " " & vbCrLf & _
    " OP. nº: " & txtReabrirOP(5) & " " & vbCrLf & _
    " Semana programada: " & txtReabrirOP(6) & " " & vbCrLf & _
    " Nome da OP.: " & txtReabrirOP(7) & " " & vbCrLf & vbCrLf & _
    "Dados do Colaborador que efetuou a baixa " & " " & vbCrLf & _
    " Chapa: " & Mid$(SkinLabel10, 1, 5) & " " & vbCrLf & _
    " Nome: " & Mid$(SkinLabel10, 9, 50) & " " & vbCrLf & _
    " Centro de Custo: " & SkinLabel12 & " " & vbCrLf & _
    " Semana baixada: " & txtReabrirOP(9) & " " & vbCrLf & _
    " Hora que realizou a baixada: " & DTPicker1.Value & " " & vbCrLf & _
    " Hora que realizou a baixada: " & txtReabrirOP(10) & " " & vbCrLf & _
    " Justificativa apresentada para reabertura da OP.: " & " " & vbCrLf & _
    " " & Combo1.Text & " "
    
    vIDCD = Format(GeraCodigoTB("tbComunicacaoDesvio", "idcd", "", ""), "000000")
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(1, 1) = vIDCD 'ID da CD
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = CDate(vDataDoBanco) 'Data de reabertura da CD
    vQualquerDado(2, 2) = "D"
    vQualquerDado(3, 1) = NomUsu 'Responsável pela abertura da CD e reabertura da Operação
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtReabrirOP(3).Text 'Nº da OS
    vQualquerDado(4, 2) = "I"
    vQualquerDado(5, 1) = vObservacao 'Observação
    vQualquerDado(5, 2) = "S"
    vQualquerDado(6, 1) = "6"
    vQualquerDado(6, 2) = "I"
    vQualquerDado(7, 1) = txtReabrirOP(13).Text 'Revisao
    vQualquerDado(7, 2) = "I"
    GravaDados "tbComunicacaoDesvio", "idcd", "I", vIDCD, 7, "", "", vIDCD
    Exit Function
Err:
    GravaCD = False
End Function

Private Function dadosEmail()
    dadosEmail = False
    Dim rsEnviaEmail As New ADODB.Recordset
    Dim SqlEnviaEmail As String
    SqlEnviaEmail = "Select email from tbUsuarios where codcoligada = '" & vCodcoligada & "' and nome = '" & NomUsu & "'"
    rsEnviaEmail.Open SqlEnviaEmail, cnBanco, adOpenKeyset, adLockOptimistic
    vEmailAprovador = rsEnviaEmail.Fields(0)
    If vEmailAprovador = "" Then
        mobjMsg.Abrir "Email do usuário LOGADO não está cadastrado", Ok, critico, "ZEUS"
        Exit Function
    End If
    rsEnviaEmail.Close
    Set rsEnviaEmail = Nothing
    dadosEmail = True
End Function

Private Sub enviaEmail()
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    vDecisao = "Aprovado"
    vCorDecisao = "#228B22"

    vSMTP = "smtp.viga.ind.br"
    vUsuEmail = "taos@viga.ind.br"
    vSenhaEmail = "taos2017@"

    With Camp
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
    End With

    With Msg
        Set .Configuration = Cof
      
'        .To = "viga@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .To = "qualidade@viga.ind.br;planejamento3@viga.ind.br;planejamento4@viga.ind.br;viga@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .From = vEmailAprovador  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "CD - Comunicação de Desvio nº: " & vIDCD
        
        .HTMLBody = "<html>" & _
        " <head>" & _
        " <meta http-equiv='Content-Type'" & _
        " content='text/html; charset=iso-8859-1'>" & _
        " <meta name='GENERATOR' content='Microsoft FrontPage Express 2.0'>" & _
        " <title>Assinatura</title>" & _
        " <STYLE type='text/css'>" & _
        " <!-- -->" & _
        " </STYLE></head>" & _
        " <body link='#0000FF' vlink='#800080'>" & _
        " <font face = 'Courier New' size = 5>" & _
        " <B><FONT STYLE='COLOR:#009ACD'> COMUNICAÇÃO DE DESVIO </FONT></B><BR><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> A CD de nº: <b>" & vIDCD & "</b>, foi aberta pelo colaborador, <b>" & NomUsu & "</b>. Onde foi detectado que: </FONT><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & vObservacao & " </FONT><BR><BR><FONT STYLE='COLOR:#009ACD'> OS nº: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & txtReabrirOP(3) & "/" & txtReabrirOP(13) & " </b><BR><FONT STYLE='COLOR:#009ACD'>Data de Abertura: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & Format(vDataDoBanco, "dd/mm/yyyy") & "</b></FONT><BR><BR>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & "" & " </FONT><BR><FONT STYLE='COLOR:#009ACD'> Att </FONT><BR><BR><BR><BR></font>" & _
        " <table border='0' cellspacing='0' width='627'>" & _
        " <tr><td width='100%'><span class='txt'>" & _
        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
        " </td></tr></table></body>"
        
        .Send
    End With
    mobjMsg.Abrir "Email enviado com suscesso!", Ok, informacao, "ZEUS"
    Exit Sub
errMail:
    Msgbox "Verifique as configurações de e-mail do usuário remetente e dos destinatários" & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "ZEUS"
    Exit Sub
End Sub


