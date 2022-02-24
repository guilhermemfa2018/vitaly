VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDevolucao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolução"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20340
   Icon            =   "frmDevolucao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   20340
   StartUpPosition =   3  'Windows Default
   Begin IMRM.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   12000
      TabIndex        =   34
      Top             =   8640
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDevolucao.frx":0CCA
      PICN            =   "frmDevolucao.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "frmDevolucao.frx":19C0
      Top             =   0
   End
   Begin VB.TextBox txtEmprestimo 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   5520
      TabIndex        =   33
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame7 
      Caption         =   "Valor Total"
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
      Left            =   9000
      TabIndex        =   31
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Digite a CHAPA do colaborador"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   "Digite parte do NOME do colaborador e tecle ENTER"
         Top             =   480
         Width           =   5655
      End
      Begin VB.CommandButton cmdDevolucao 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   8
         ToolTipText     =   "Realiza a pesquisa de todos os colaboradores"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   8400
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   125960193
         CurrentDate     =   42632
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   8400
         OleObjectBlob   =   "frmDevolucao.frx":1BF4
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "frmDevolucao.frx":1C6A
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDevolucao.frx":1CCC
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Setor "
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
      Left            =   3360
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   2040
         TabIndex        =   28
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDevolucao.frx":1D30
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmDevolucao.frx":1D9C
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Função"
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
      Left            =   3360
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDevolucao.frx":1E04
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmDevolucao.frx":1E6A
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Valor Total"
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
      Left            =   17280
      TabIndex        =   19
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Text            =   "0"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Informações "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10320
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDevolucao.frx":1ECC
         TabIndex        =   2
         Top             =   720
         Width           =   5415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDevolucao.frx":1F94
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmDevolucao.frx":2048
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin IMRM.chameleonButton cmdDev 
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   17
      Top             =   8640
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
      MICON           =   "frmDevolucao.frx":20DC
      PICN            =   "frmDevolucao.frx":20F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin IMRM.chameleonButton cmdDev 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   8640
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
      MICON           =   "frmDevolucao.frx":2DD2
      PICN            =   "frmDevolucao.frx":2DEE
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
      Caption         =   "Itens a serem devolvidos "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   20055
      Begin VB.CommandButton cmdDevolucao 
         Caption         =   ">"
         Height          =   615
         Index           =   1
         Left            =   11880
         TabIndex        =   15
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdDevolucao 
         Caption         =   "<"
         Height          =   615
         Index           =   2
         Left            =   11880
         TabIndex        =   14
         Top             =   3600
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6975
         Left            =   12600
         TabIndex        =   16
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   12303
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6975
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12303
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
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
End
Attribute VB_Name = "frmDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPosicaoAtual As Integer
Private rsCriterio As New ADODB.Recordset
Private SqlCriterio As String
Private vIDMovEmprestimo As Double
Private vCodColigadaEmprestimo As Integer

Private Sub chameleonButton1_Click()
    FCREmprestimo.Show 1
End Sub

Private Sub cmdDev_Click(Index As Integer)
    Select Case Index
    Case 0
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "Ferramentaria"
            varGlobal = txtEmprestimo(0).Text
            If ListView1.ListItems.Count = 0 Then
                mobjMsg.Abrir "O Colaborador não possui nenhuma ferramenta emprestada em seu nome", Ok, informacao, "Ferramentaria"
                Unload Me
            Else
                FCREmprestimo.Show 1
            End If
            Unload Me
        Else
            SkinLabel1.Visible = False
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "Ferramentaria"
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub cmdDevolucao_Click(Index As Integer)
    Select Case Index
    Case 1
        addRemLoteNota ListView1, ListView2
        vQtdSolicitada = 0
        SomaLV ListView1, 11, Text2
        SomaLV ListView2, 11, Text1
    Case 2
        addRemLoteNota ListView2, ListView1
        SomaLV ListView1, 11, Text2
        SomaLV ListView2, 11, Text1
    End Select
End Sub

Private Sub Form_Activate()
On Error GoTo Err
    ListView1.SetFocus
    Exit Sub
Err:
    If Err.Number = 5 Then Resume Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "Ferramentaria"
            varGlobal = txtEmprestimo(0).Text
            If ListView1.ListItems.Count = 0 Then
                mobjMsg.Abrir "O Colaborador não possui nenhuma ferramenta emprestada em seu nome", Ok, informacao, "Ferramentaria"
                Unload Me
            Else
                FCREmprestimo.Show 1
            End If
        End If
    ElseIf KeyCode = 118 Then
        Unload Me
    End If
End Sub

Private Sub ListView1_Click()
    MarcaDesmarca ListView1
End Sub

Private Sub ListView1_DblClick()
    addRemLoteNota ListView1, ListView2
    vQtdSolicitada = 0
    DoEvents
    SomaLV ListView1, 11, Text2
    SomaLV ListView2, 11, Text1
    ListView1.SetFocus
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    MarcaDesmarca ListView1
    If KeyCode = 13 Or KeyCode = 9 Or KeyCode = 32 Then ' Enter ou TAB
        addRemLoteNota ListView1, ListView2
        vQtdSolicitada = 0
        DoEvents
        SomaLV ListView1, 11, Text2
        SomaLV ListView2, 11, Text1
        ListView1.SetFocus
    End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    MarcaDesmarca ListView1
End Sub

Private Sub ListView2_Click()
    MarcaDesmarca ListView2
End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    MarcaDesmarca ListView2
    If KeyCode = 13 Or KeyCode = 9 Or KeyCode = 32 Then ' Enter ou TAB
        addRemLoteNota ListView2, ListView1
        DoEvents
        SomaLV ListView1, 11, Text2
        SomaLV ListView2, 11, Text1
        ListView2.SetFocus
    End If
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
    MarcaDesmarca ListView2
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    vCodLocalEstoque = vLocalEstoque
    listview_cabecalho

    chamaSQL "select b.localestoque as codloc,a.idmov,b.codigoprd,(b.qtdemprestado-b.qtddevolvida) as QTDEPENDENTE,b.um,b.descricao,a.dataemprestimo as DATAEMISSAO,qtDiasEmp = p.CAMPOLIVRE ,dife = CONVERT(VARCHAR, DATEDIFF(DAY, M.DATAEMISSAO ,GETDATE()) ) ,manutencao = ( case when ( SELECT manu.DATAVENCIMENTO from " & vBancoSAP & ".dbo.OFVENCPLANOMANUT manu INNER join " & vBancoSAP & ".dbo.TPRODUTO Prd on manu.IDOBJOF = SUBSTRING(Prd.CODIGOPRD,4,9) AND PRD.CODIGOPRD  = P.CODIGOPRD) < GETDATE()then 'Sim' else 'Não' end),recolher  =  case when (p.CAMPOLIVRE - CONVERT(VARCHAR, DATEDIFF(DAY, a.dataemprestimo ,GETDATE()) )) <0 then  'Sim' else 'Não' end, " & _
             "b.valortotal/b.qtdemprestado as valor_unit,b.idprd from tbEmprestimo as a inner join tbEmprestimoItens as b on a.idmov = b.idmov and (b.qtdemprestado-b.qtddevolvida) > 0 inner join " & vBancoSAP & ".dbo.tloc as c on b.localestoque = c.CODLOC COLLATE SQL_Latin1_General_CP1_CI_AS and c.CODFILIAL = 1 inner join " & vBancoSAP & ".dbo.TMOV as m on a.codcoligada = m.CODCOLIGADA and CAST(a.numeromov AS INT) = m.NUMEROMOV COLLATE SQL_Latin1_General_CP1_CI_AS and a.serie = m.SERIE COLLATE SQL_Latin1_General_CP1_CI_AS and a.idmov = m.IDMOV and m.CODFILIAL = 1 inner join " & vBancoSAP & ".dbo.TPRODUTO P on b.idprd = p.IDPRD where a.codcoligada = 1 and a.chapa = '" & Mid$(varGlobal, 1, 6) & "' AND A.localestoque = " & Val(vLocalEstoque) & ""
    Compoe_Listview ListView1, Sqlp, "00"
ResultPesq
    PersonaColLVForm ListView1, 9, "N", "S", "", "N", "N", "S", "D"
    PersonaColLVForm ListView1, 10, "N", "S", "", "N", "N", "S", "D"
    PersonaColLVForm ListView1, 11, "N", "N", "", "N", "N", "S", "D"

    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    Me.Top = (Principal.Height / 2) - (Me.Height / 2)

'===============================
    AplicarSkin Me, Principal.Skin1
    Text1 = "Nenhum"

    On Error Resume Next

    Option1 = True
    Check1 = 1
    'Recupera o nome do Skin atual
    Dim P_Buffer As String
    Dim P_M() As String

    P_Buffer = GetProfileSection("Skin", App.Path & "\CONFIG.INI")
    P_M = Split(P_Buffer, vbNullChar)
    SkinLabel2.Caption = Join(P_M, vbCrLf)
    
    Skin1.LoadSkin App.Path & "\skins\Effective.skn"
    Skin1.ApplySkin Me.hwnd
'===============================
    
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "LOC", ListView1.Width / 15
    ListView1.ColumnHeaders.Add , , "ID MOV", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "COD PRD", ListView1.Width / 8.8
    ListView1.ColumnHeaders.Add , , "QTDE", ListView1.Width / 17.5
    ListView1.ColumnHeaders.Add , , "UN", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "DESCRIÇÃO", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "DATA EMP", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "DIAS EMP", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "DIAS OBRA", ListView1.Width / 10.5
    ListView1.ColumnHeaders.Add , , "MNT P", ListView1.Width / 15.5
    ListView1.ColumnHeaders.Add , , "RECOLHE", ListView1.Width / 13
    ListView1.ColumnHeaders.Add , , "VALOR", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "ID PRD", ListView1.Width / 10000
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "LOC", ListView2.Width / 10
    ListView2.ColumnHeaders.Add , , "ID MOV", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "COD PRD", ListView2.Width / 5.5
    ListView2.ColumnHeaders.Add , , "QTDE", ListView2.Width / 12
    ListView2.ColumnHeaders.Add , , "UN", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "DESCRIÇÃO", ListView2.Width / 2.5
    ListView2.ColumnHeaders.Add , , "DATA EMP", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "DIAS EMP", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "DIAS OBRA", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "MNT P", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "RECOLHE", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "VALOR", ListView2.Width / 9
    ListView2.ColumnHeaders.Add , , "ID PRD", ListView2.Width / 10000
    Me.ListView2.ColumnHeaders(12).Alignment = lvwColumnRight

    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    
    ListView1.BackColor = RGB(135, 194, 194)
    ListView2.BackColor = RGB(135, 194, 194)
    
End Sub

Private Sub ResultPesq()
    SqlCriterio = "Select a.chapa,a.nome,dataemprestimo,a.numeromov,a.nomequememprestou,A.codfuncao,A.nomefuncao,A.codsecao,A.nomesecao, a.codcoligada from tbEmprestimo as a where codcoligada = 1 and a.localestoque ='" & Val(vLocalEstoque) & "' and a.chapa = '" & Mid$(varGlobal, 1, 6) & "' and a.numeromov = '" & Mid$(varGlobal, 7, 6) & "'"
    rsCriterio.Open SqlCriterio, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCriterio.RecordCount = 0 Then
        rsCriterio.Close
        Set rsCriterio = Nothing
        SqlCriterio = "Select a.chapa,a.nome,dataemprestimo,a.numeromov,a.nomequememprestou,A.codfuncao,A.nomefuncao,A.codsecao,A.nomesecao, a.codcoligada from tbEmprestimo as a where codcoligada = 1 and a.localestoque ='" & Val(vLocalEstoque) & "' and a.chapa = '" & Mid$(varGlobal, 1, 6) & "'"
        rsCriterio.Open SqlCriterio, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    If rsCriterio.RecordCount > 0 Then
        compoeControlesForm
    End If
    rsCriterio.Close
    Set rsCriterio = Nothing
End Sub

Private Sub compoeControlesForm()
    txtEmprestimo(0) = rsCriterio.Fields(0) 'Chapa do colaborador
    txtEmprestimo(1) = rsCriterio.Fields(1) 'Nome do colaborador
    DTPicker1.Value = Date 'rsCriterio.Fields(2) 'Data do emprestimo
    
    txtEmprestimo(2) = rsCriterio.Fields(5) 'Código da Função
    txtEmprestimo(3) = rsCriterio.Fields(6) 'Nome da Função
    txtEmprestimo(4) = rsCriterio.Fields(7) 'Código da Seção
    txtEmprestimo(5) = rsCriterio.Fields(8) 'nome da Seção
    vCodColigadaEmprestimo = rsCriterio.Fields(9) 'Coidigo da coligada que realizou o emprestimo
    
'    txtEmprestimo(8) = rsCriterio.Fields(4) 'Nome do colaborador emprestou o produto
End Sub

Private Sub MarcaDesmarca(LV As ListView)
    'Deixa checado somente um item do Listview
    Dim X As Integer, Y As Integer, J As Integer
    Y = LV.ListItems.Count
    If Y = 0 Then Exit Sub
    J = LV.SelectedItem.Index
    For X = 1 To Y
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        End If
    Next
    LV.ListItems.Item(J).Checked = True
    vPosicaoAtual = J
End Sub

Private Sub addRemLoteNota(lvOrigem As ListView, lvDestino As ListView)
On Error GoTo Err
    Dim X As Integer, Y As Integer, X1 As Integer, Y1 As Integer
    Dim ItemLst As ListItem
    Y = lvOrigem.ListItems.Count
    For X = 1 To Y
        If Y < X Then
            'Exit Sub
            Exit For
        End If
        lvOrigem.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        
        If lvOrigem.ListItems(X).Checked = True Then
            If lvOrigem.Name = "ListView1" Then

                If lvOrigem.SelectedItem.ListSubItems.Item(3) > 1 Then
                    vQtdDisponivel = lvOrigem.SelectedItem.ListSubItems.Item(3)
                    frmInformaQtd.Show 1
                    If vQtdSolicitada = 0 Then Exit Sub
                Else
                    vQtdDisponivel = 1
                    vQtdSolicitada = 1
                End If
                
                Y1 = lvDestino.ListItems.Count
                
                'VERIFICA SE O PRODUTO JÁ SE ENCONTRA NA LISTA DE EMPRESTIMO
                For X1 = 1 To Y1
                    If Y1 < X1 Then
                        'Exit Sub
                        Exit For
                    End If
                    lvDestino.ListItems.Item(X1).Selected = True 'Passar a selecao para o próximo item
                    If lvOrigem.SelectedItem.ListSubItems.Item(2) = lvDestino.SelectedItem.ListSubItems.Item(2) Then
                        lvDestino.SelectedItem.ListSubItems.Item(3) = Val(lvDestino.SelectedItem.ListSubItems.Item(3)) + vQtdSolicitada
                        lvDestino.SelectedItem.ListSubItems.Item(11) = Format(lvDestino.SelectedItem.ListSubItems.Item(11) * lvDestino.SelectedItem.ListSubItems.Item(3), "#,##0.00;(#,##0.00)") 'Valor total
                        If vQtdSolicitada = Val(lvOrigem.SelectedItem.ListSubItems.Item(3)) Then
                            lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                        Else
                            lvOrigem.SelectedItem.ListSubItems.Item(3) = Val(lvOrigem.SelectedItem.ListSubItems.Item(3)) - vQtdSolicitada
                        End If
                        Exit Sub
                    End If
                Next
            ElseIf lvOrigem.Name = "ListView2" Then
                Y1 = lvDestino.ListItems.Count
                For X1 = 1 To Y1
                    If Y1 < X1 Then
                        lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                        Exit Sub
                    End If
                    lvDestino.ListItems.Item(X1).Selected = True 'Passar a selecao para o próximo item
                    If lvOrigem.SelectedItem.ListSubItems.Item(2) = lvDestino.SelectedItem.ListSubItems.Item(2) Then
                        lvDestino.SelectedItem.ListSubItems.Item(3) = Val(lvDestino.SelectedItem.ListSubItems.Item(3)) + Val(lvOrigem.SelectedItem.ListSubItems.Item(3))
                        lvDestino.SelectedItem.ListSubItems.Item(11) = Format(lvDestino.SelectedItem.ListSubItems.Item(11) * lvDestino.SelectedItem.ListSubItems.Item(3), "#,##0.00;(#,##0.00)") 'Valor total
                        
                        lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                        Exit Sub
                    End If
                Next
            End If
            
            Set ItemLst = lvDestino.ListItems.Add(, , lvOrigem.ListItems(X)) ' Local de estoque
            ItemLst.SubItems(1) = "" & lvOrigem.SelectedItem.ListSubItems.Item(1) 'Código do MOVIMENTO
            ItemLst.SubItems(2) = "" & lvOrigem.SelectedItem.ListSubItems.Item(2) 'Código do produto
            If lvOrigem.Name = "ListView1" Then
                If lvOrigem.SelectedItem.ListSubItems.Item(3) > 1 Then
                    ItemLst.SubItems(3) = "" & vQtdSolicitada 'Quantidade
                Else
                    ItemLst.SubItems(3) = "" & lvOrigem.SelectedItem.ListSubItems.Item(3) 'Quantidade
                End If
            Else
                ItemLst.SubItems(3) = "" & lvOrigem.SelectedItem.ListSubItems.Item(3) 'Quantidade
            End If
            ItemLst.SubItems(4) = "" & lvOrigem.SelectedItem.ListSubItems.Item(4) 'Unidade de medida
            ItemLst.SubItems(5) = "" & lvOrigem.SelectedItem.ListSubItems.Item(5) 'Descrição do produto
            ItemLst.SubItems(6) = "" & lvOrigem.SelectedItem.ListSubItems.Item(6) 'DATA DE EMPRESTIMO DA FERAMENTA
            ItemLst.SubItems(7) = "" & lvOrigem.SelectedItem.ListSubItems.Item(7) 'DIAS EMPRESTIMO
            ItemLst.SubItems(8) = "" & lvOrigem.SelectedItem.ListSubItems.Item(8) 'DIAS OBRA
            ItemLst.SubItems(9) = "" & lvOrigem.SelectedItem.ListSubItems.Item(9) 'MANUTENÇÃO PREVENTIVA VENCENDO?
            ItemLst.SubItems(10) = "" & lvOrigem.SelectedItem.ListSubItems.Item(10) 'RECOLHER?
            ItemLst.SubItems(11) = "" & Format(lvOrigem.SelectedItem.ListSubItems.Item(11), "#,##0.00;(#,##0.00)") 'VALOR
            ItemLst.SubItems(12) = "" & lvOrigem.SelectedItem.ListSubItems.Item(12) 'IDENTIFICADOR DO PRODUTO
            If lvOrigem.Name = "ListView1" Then
                If vQtdDisponivel = vQtdSolicitada Then
                    lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                Else
                    lvOrigem.SelectedItem.ListSubItems.Item(3) = lvOrigem.SelectedItem.ListSubItems.Item(3) - vQtdSolicitada
                    lvOrigem.ListItems(X).Checked = False
                End If
            ElseIf lvOrigem.Name = "ListView2" Then
                If Y1 < X1 Then
                    lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                End If
            End If
            Y = Y - 1
            X = X - 1
        End If
    Next
    lvOrigem.ListItems.Item(vPosicaoAtual).Selected = True 'Passar a selecao para o próximo item
    
    'Ordena listview para exibir na tela
    lvDestino.Sorted = True
    lvDestino.SortKey = 4
    lvDestino.SortOrder = lvwAscending
    lvDestino.Refresh
    lvOrigem.Sorted = True
    lvOrigem.SortKey = 4
    lvOrigem.SortOrder = lvwAscending
    lvOrigem.Refresh
    Exit Sub
Err:
    If Err.Number = 35600 Then
        Exit Sub
    End If
End Sub


'---------------------------- ROTINAS DE GRAVAÇÃO DA DEVOLUÇÃO --------------------------------
'----------------------------------------------------------------------------------------------

Private Function salvar_Dados()
'On Error GoTo Err
    If ValidaCampo = False Then Exit Function
    vTransacaoAtiva = 1
    cnBanco.BeginTrans
    salvar_Dados = True
    
    AtualizaEmprestimo
    
    GeraNumeroMov
    GeraIDMov
    GeraSequencialEstoque
    
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    limpaQualquerDado
    vQualquerDado(1, 1) = txtEmprestimo(0).Text 'Identificador do colaborador que está devolvendo a ferramenta
    vQualquerDado(1, 2) = "S"
    vQualquerDado(2, 1) = txtEmprestimo(1) 'Nome do colaborador que está devolvendo a ferramenta
    vQualquerDado(2, 2) = "S"
    vQualquerDado(3, 1) = txtEmprestimo(2).Text ' Identificador da função do colaborador que está devolvendo a ferramenta
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtEmprestimo(3).Text ' Nome da função do colaborador que está devolvendo a ferramenta
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtEmprestimo(4).Text ' Identificador da setor do colaborador que está devolvendo a ferramenta
    vQualquerDado(5, 2) = "S"
    vQualquerDado(6, 1) = txtEmprestimo(5).Text ' Nome da setor do colaborador que está devolvendo a ferramenta
    vQualquerDado(6, 2) = "S"
    
    vQualquerDado(7, 1) = vIDMov ' VIDMOV - Identificador do movimento gerado pelo sistema deferramentaria
    vQualquerDado(7, 2) = "I"
    
    vQualquerDado(8, 1) = Format(vNumeromov, "000000") ' VNUMEROMOV - Numero do movimento gerado pelo sistema deferramentaria
    vQualquerDado(8, 2) = "S"
    
    vQualquerDado(9, 1) = vSerie ' SERIE - Serie do movimento da ferramentaria
    vQualquerDado(9, 2) = "S"
    
    vQualquerDado(10, 1) = 1 ' Código da Coligada
    vQualquerDado(10, 2) = "I"
    
    vQualquerDado(11, 1) = vCodLocalEstoque ' Local de estoque
    vQualquerDado(11, 2) = "I"
    
    vQualquerDado(12, 1) = vCodVenRM & " - " & vNomeVenRM 'NomUsu ' Nome de quem está recebendo a ferramenta
    vQualquerDado(12, 2) = "S"
    
    vQualquerDado(13, 1) = vCodUsuarioRM ' codusuario (RM) de quem devolveu
    vQualquerDado(13, 2) = "S"
    
    txtEmprestimo(9) = vNumeromov
    GravaDados "tbDevolucao", "numeromov", "S", txtEmprestimo(9), 13, "", "", txtEmprestimo(9)
    cnBanco.CommitTrans
        
    
    limpaQualquerDado
    GravaProdutosDevolvidos

    vTransacaoAtiva = 0
    Exit Function
Err:
    cnBanco.RollbackTrans
    salvar_Dados = False
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    
    If ListView2.ListItems.Count = 0 Then
        mobjMsg.Abrir "Nenhuma ferramenta foi devolvida", Ok, critico, "Atenção"
        ListView1.SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub AtualizaEmprestimo()
    Dim rsAtualizaEmprestimo As New ADODB.Recordset
    Dim SqlAtualizaEmprestimo As String
    Dim X As Integer, Y As Integer, X1 As Integer, Y1 As Integer, vContaPrd As Integer
    If Val(Text2.Text) = 0 Then
        Y = ListView2.ListItems.Count
        'TABELA: TBEMPRESTIMO
        'ALTERA O STATUS PARA 'D' - DEVOLVIDO NA TABELA DE EMPRESTIMO
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
            SqlAtualizaEmprestimo = "update tbEmprestimo set status = 'D' where codcoligada = '" & vCodColigadaRM & "' and localestoque = '" & Val(vCodLocalEstoque) & "' and idmov = '" & ListView2.SelectedItem.ListSubItems.Item(1) & "'"
            rsAtualizaEmprestimo.Open SqlAtualizaEmprestimo, cnBanco
            Set rsAtualizaEmprestimo = Nothing
        Next
        
        'TABELA TBEMPRESTIMOITENS
        'ALTERA O STATUS DOS ITENS PARA 'D' - DEVOLVIDO NA TABELA DE ITENS EMPRESTADOS
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
            SqlAtualizaEmprestimo = "update tbEmprestimoItens set status = 'D', qtddevolvida = qtddevolvida + '" & ListView2.SelectedItem.ListSubItems.Item(3) & "' where codcoligada = '" & vCodColigadaRM & "' and localestoque = '" & vCodLocalEstoque & "' and idmov = '" & ListView2.SelectedItem.ListSubItems.Item(1) & "' and idprd = '" & ListView2.SelectedItem.ListSubItems.Item(12) & "'"
            rsAtualizaEmprestimo.Open SqlAtualizaEmprestimo, cnBanco
            Set rsAtualizaEmprestimo = Nothing
        Next
    Else
        Y = ListView2.ListItems.Count
        Y1 = ListView1.ListItems.Count
        'TABELA: TBEMPRESTIMO
        'ALTERA O STATUS PARA 'D' - DEVOLVIDO NA TABELA DE EMPRESTIMO. SE A QUANTIDADE DEVOLVIDA FOR MENOR QUE A QUANTIDADE EMPRESTADA
        'O STATUS NÃO MUDA
        
        If Y1 = 0 Then
            SqlAtualizaEmprestimo = "update tbEmprestimo set status = 'D' where codcoligada = '" & vCodColigadaRM & "' and localestoque = '" & Val(vCodLocalEstoque) & "' and idmov = '" & ListView2.SelectedItem.ListSubItems.Item(1) & "'"
            rsAtualizaEmprestimo.Open SqlAtualizaEmprestimo, cnBanco
            Set rsAtualizaEmprestimo = Nothing
        End If
        
        'TABELA TBEMPRESTIMOITENS
        'ALTERA O STATUS DOS ITENS PARA 'D' - DEVOLVIDO NA TABELA DE ITENS EMPRESTADOS. SE A QUANTIDADE DEVOLVIDA FOR MENOR QUE A QUANTIDADE EMPRESTADA
        'O STATUS NÃO MUDA
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
            vContaPrd = 0
            For X1 = 1 To Y1
                ListView1.ListItems.Item(X1).Selected = True 'Passar a selecao para o próximo item
                If ListView1.SelectedItem.ListSubItems.Item(2) = ListView2.SelectedItem.ListSubItems.Item(2) Then
                    vContaPrd = vContaPrd + 1
                End If
            Next
            If vContaPrd = 0 Then
                SqlAtualizaEmprestimo = "update tbEmprestimoItens set status = 'D', qtddevolvida = qtddevolvida + '" & ListView2.SelectedItem.ListSubItems.Item(3) & "' where codcoligada = '" & vCodColigadaRM & "' and localestoque = '" & vCodLocalEstoque & "' and idmov = '" & ListView2.SelectedItem.ListSubItems.Item(1) & "' and idprd = '" & ListView2.SelectedItem.ListSubItems.Item(12) & "'"
                rsAtualizaEmprestimo.Open SqlAtualizaEmprestimo, cnBanco
                Set rsAtualizaEmprestimo = Nothing
            Else
                SqlAtualizaEmprestimo = "update tbEmprestimoItens set status = 'E', qtddevolvida = qtddevolvida + '" & ListView2.SelectedItem.ListSubItems.Item(3) & "' where codcoligada = '" & vCodColigadaRM & "' and localestoque = '" & vCodLocalEstoque & "' and idmov = '" & ListView2.SelectedItem.ListSubItems.Item(1) & "' and idprd = '" & ListView2.SelectedItem.ListSubItems.Item(12) & "'"
                rsAtualizaEmprestimo.Open SqlAtualizaEmprestimo, cnBanco
                Set rsAtualizaEmprestimo = Nothing
            End If
        Next
    End If
End Sub

Private Function GeraNumeroMov()
    Dim rsGeraNumeroMov As New ADODB.Recordset
    Dim SqlGeraNumeroMov As String
    
    SqlGeraNumeroMov = "Select top 1 * from tbMov as a order by numeromov Desc"
    rsGeraNumeroMov.Open SqlGeraNumeroMov, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraNumeroMov.RecordCount > 0 Then
        vNumeromov = Val(Mid$(rsGeraNumeroMov.Fields(0), 1, 6)) + 1
    Else
        vNumeromov = 1
    End If
    vNumeromov = Format(vNumeromov, "000000")
    vSerie = SerieEmpresa & "D" & vLocalEstoque
    rsGeraNumeroMov.Close
    Set rsGeraNumeroMov = Nothing
    limpaQualquerDado
    
    vQualquerDado(1, 1) = vNumeromov ' VNUMEROMOV - Numero do movimento para DEVOLUÇÃO gerado pelo sistema de ferramentaria
    vQualquerDado(1, 2) = "S"
    vQualquerDado(2, 1) = vSerie ' SERIE - Serie do movimento de DEVOLUÇÃO da ferramentaria
    vQualquerDado(2, 2) = "S"
    vQualquerDado(3, 1) = 1 ' Código da Coligada
    vQualquerDado(3, 2) = "I"
    'GRAVA DADOS DA DEVOLUÇÃO
    GravaDados "tbMov", "Numeromov", "S", txtEmprestimo(0), 3, "", "", txtEmprestimo(0)
End Function

Private Function GeraIDMov()
    Dim rsGeraIDMov As New ADODB.Recordset
    Dim SqlGeraIDMov As String
    
    Dim rsAtualizaIDMov As New ADODB.Recordset
    Dim SqlAtualizaIDMov As String
    
    SqlGeraIDMov = "select * from " & vBancoSAP & ".dbo.GAUTOINC as a where a.codautoinc like 'IDMOV' and a.codcoligada = '" & vCodColigadaRM & "'"
    
    rsGeraIDMov.Open SqlGeraIDMov, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraIDMov.RecordCount > 0 Then
        vIDMov = Val(rsGeraIDMov.Fields(3)) + 1
    Else
        vIDMov = 1
    End If
    rsGeraIDMov.Close
    Set rsGeraIDMov = Nothing

    SqlAtualizaIDMov = "UPDATE GAUTOINC set VALAUTOINC = " & vIDMov & " where codautoinc like 'IDMOV' and codcoligada = '" & vCodColigadaRM & "'"
    'rsAtualizaIDMov.Open SqlAtualizaIDMov, cnBancoSAP
    Set rsAtualizaIDMov = cnBancoSAP.Execute(SqlAtualizaIDMov)
    Set rsAtualizaIDMov = Nothing

End Function

Private Function GeraSequencialEstoque()
    Dim rsSequencialEstoque As New ADODB.Recordset
    Dim SqlSequencialEstoque As String
    
    Dim rsAtuSequencialEstoque As New ADODB.Recordset
    Dim SqlAtuSequencialEstoque As String
    
    SqlSequencialEstoque = "select * from " & vBancoSAP & ".dbo.GAUTOINC as a where a.codautoinc like 'SEQUENCIALESTOQUE' and a.codcoligada = '" & vCodColigadaRM & "'"
    
    rsSequencialEstoque.Open SqlSequencialEstoque, cnBanco, adOpenKeyset, adLockReadOnly
    If rsSequencialEstoque.RecordCount > 0 Then
        vSequencialEstoque = Val(rsSequencialEstoque.Fields(3)) + 1
    Else
        vSequencialEstoque = 1
    End If
    rsSequencialEstoque.Close
    Set rsSequencialEstoque = Nothing
    
    SqlAtuSequencialEstoque = "UPDATE GAUTOINC set VALAUTOINC = " & vSequencialEstoque & " where codautoinc like 'SEQUENCIALESTOQUE' and codcoligada = '" & vCodColigadaRM & "'"
    rsAtuSequencialEstoque.Open SqlAtuSequencialEstoque, cnBancoSAP
    Set rsAtuSequencialEstoque = Nothing

End Function

Private Sub GravaProdutosDevolvidos()
On Error GoTo Err
    cnBanco.BeginTrans
    
    Dim rsGravaProdutosDevolvidos As New ADODB.Recordset
    Dim sqlGravaProdutosDevolvidos As String
    Dim X As Integer, Y As Integer
    
    sqlGravaProdutosDevolvidos = "Select * from tbDevolucaoItens as a where a.codcoligada = '" & vCodColigadaRM & "' and a.numeromov = '" & vNumeromov & "'"
    rsGravaProdutosDevolvidos.Open sqlGravaProdutosDevolvidos, cnBanco, adOpenKeyset, adLockOptimistic
    
    ListView2.ListItems.Item(1).Selected = True
    
    GravaTMov 'grava dados na tabela TMOV (TOTVS RM)
'    GravaRelacEntreMovCab
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaProdutosDevolvidos.AddNew
        vIDMovEmprestimo = ListView2.SelectedItem.ListSubItems.Item(1)
        If X = 1 Then
            GravaRelacEntreMovCab
        End If
        rsGravaProdutosDevolvidos(0) = Format(vNumeromov, "000000") 'Numero do Movimento
        rsGravaProdutosDevolvidos(1) = vCodColigadaRM 'coligada
        rsGravaProdutosDevolvidos(2) = X 'Sequencial
        rsGravaProdutosDevolvidos(3) = ListView2.ListItems.Item(X) 'Local de estoque
        rsGravaProdutosDevolvidos(4) = ListView2.SelectedItem.ListSubItems.Item(2) 'Código do produto
        rsGravaProdutosDevolvidos(5) = ListView2.SelectedItem.ListSubItems.Item(5) 'Descrição do produto
        rsGravaProdutosDevolvidos(6) = vIDMov 'Identificador do movimento
        rsGravaProdutosDevolvidos(7) = ListView2.SelectedItem.ListSubItems.Item(12) 'Identificador do produto
        rsGravaProdutosDevolvidos(8) = ListView2.SelectedItem.ListSubItems.Item(3) 'quantidade devolvida
        rsGravaProdutosDevolvidos(9) = DTPicker1.Value 'Data da devolucao
        rsGravaProdutosDevolvidos(10) = Time 'Hora do empréstimo
        rsGravaProdutosDevolvidos(11) = NomUsu 'Nome de quem emprestou as ferramentas
        rsGravaProdutosDevolvidos(12) = ListView2.SelectedItem.ListSubItems.Item(4) 'Unidade de medida do produto
        rsGravaProdutosDevolvidos(13) = ListView2.SelectedItem.ListSubItems.Item(11) 'Valor total
        rsGravaProdutosDevolvidos(14) = vSerie 'Serie do movimento
        rsGravaProdutosDevolvidos(15) = ListView2.SelectedItem.ListSubItems.Item(1) 'Identificador do movimento de emprestimo
        
        GravaTitMMov X 'grava dados na tabela TITMMOV (TOTVS RM)
        GravaTprdLoc 'grava dados na tabela TPRDLOC (TOTVS RM)
        
        GravaRelacEntreMovItens ListView2.SelectedItem.ListSubItems.Item(1), ListView2.SelectedItem.ListSubItems.Item(12), X, ListView2.SelectedItem.ListSubItems.Item(3)
        
    Next
    If Not rsGravaProdutosDevolvidos.EOF Then rsGravaProdutosDevolvidos.Update
    rsGravaProdutosDevolvidos.Close
    cnBanco.CommitTrans
    Exit Sub
Err:
    cnBanco.RollbackTrans
End Sub

Private Sub GravaTMov()
On Error GoTo Err
    cnBancoSAP.BeginTrans
    
    Dim rsGravaTMov As New ADODB.Recordset
    Dim SqlGravaTMov As String
    Dim vValor As Double
   
    SqlGravaTMov = "Select A.CODCOLIGADA,A.IDMOV,A.CODFILIAL,A.CODLOC,A.CODCFO,A.CODCFONATUREZA,A.NUMEROMOV,A.SERIE,A.CODTMV,A.TIPO,A.STATUS,A.MOVIMPRESSO,A.DOCIMPRESSO,A.FATIMPRESSA,A.DATAEMISSAO,A.COMISSAOREPRES,A.VALORBRUTO,A.VALORLIQUIDO,A.VALOROUTROS,A.PERCCOMISSAO,A.PESOLIQUIDO," & _
    "A.PESOBRUTO,A.CODMOEVALORLIQUIDO,A.DATAMOVIMENTO,A.GEROUFATURA,A.CODCFOAUX,A.CODVEN1,A.CODVEN2,A.PERCCOMISSAOVEN2,A.CODCOLCFO,A.CODCOLCFONATUREZA,A.CODUSUARIO,A.GERADOPORLOTE,A.STATUSEXPORTCONT,A.GEROUCONTATRABALHO,A.GERADOPORCONTATRABALHO,A.HORULTIMAALTERACAO," & _
    "A.INDUSOOBJ,A.CONTABILIZADOPORTOTAL,A.INTEGRADOBONUM,A.FLAGPROCESSADO,A.ABATIMENTOICMS,A.USUARIOCRIACAO,A.DATACRIACAO,A.STSEMAIL,A.VALORBRUTOINTERNO,A.VINCULADOESTOQUEFL,A.VALORDESCCONDICIONAL,A.VALORDESPCONDICIONAL,A.CONTORCAMENTOANTIGO,A.SEQUENCIALESTOQUE," & _
    "A.INTEGRADOAUTOMACAO,A.INTEGRAAPLICACAO,A.DATALANCAMENTO,A.EXTENPORANEO,A.RECIBONFESTATUS,A.IDMOVCFO,A.VALORMERCADORIAS,A.USARATEIOVALORFIN,A.CODCOLCFOAUX,A.VRBASEINSSOUTRAEMPRESA,A.VALORBRUTOORIG,A.VALORLIQUIDOORIG,A.VALOROUTROSORIG,A.RECCREATEDBY,A.RECCREATEDON," & _
    "A.RECMODIFIEDBY,A.RECMODIFIEDON from tmov as a where a.idmov = '" & vIDMov & "'"
    rsGravaTMov.Open SqlGravaTMov, cnBancoSAP, adOpenKeyset, adLockOptimistic
 
    If Text1.Text = "Nenhum" Then
        vValor = Format(0, "#,##0.00;(#,##0.00)")
    Else
        vValor = Format(Text1.Text, "#,##0.00;(#,##0.00)")
    End If
    
    If rsGravaTMov.RecordCount = 0 Then
        rsGravaTMov.AddNew
        rsGravaTMov.Fields(0) = vCodColigadaRM
        rsGravaTMov.Fields(1) = vIDMov
        rsGravaTMov.Fields(2) = 1
        rsGravaTMov.Fields(3) = vCodLocalEstoque
        rsGravaTMov.Fields(4) = "000001"
        rsGravaTMov.Fields(5) = "000001"
        rsGravaTMov.Fields(6) = Format(vNumeromov, "000000")
        rsGravaTMov.Fields(7) = vSerie
        rsGravaTMov.Fields(8) = "1.2.16"
        rsGravaTMov.Fields(9) = "P"
        rsGravaTMov.Fields(10) = "N"
        rsGravaTMov.Fields(11) = 0
        rsGravaTMov.Fields(12) = 0
        rsGravaTMov.Fields(13) = 0
        rsGravaTMov.Fields(14) = DTPicker1.Value
        rsGravaTMov.Fields(15) = 0
        rsGravaTMov.Fields(16) = vValor
        rsGravaTMov.Fields(17) = vValor
        rsGravaTMov.Fields(18) = vValor
        rsGravaTMov.Fields(19) = 0
        rsGravaTMov.Fields(20) = 0
        rsGravaTMov.Fields(21) = 0
        rsGravaTMov.Fields(22) = "R$"
        rsGravaTMov.Fields(23) = DTPicker1.Value
        rsGravaTMov.Fields(24) = 0
        rsGravaTMov.Fields(25) = "CXXXXXXXXXX"
        rsGravaTMov.Fields(26) = txtEmprestimo(0).Text
        rsGravaTMov.Fields(27) = vCodVenRM
        rsGravaTMov.Fields(28) = 0
        rsGravaTMov.Fields(29) = 1
        rsGravaTMov.Fields(30) = 1
        rsGravaTMov.Fields(31) = vCodUsuarioRM
        rsGravaTMov.Fields(32) = 0
        rsGravaTMov.Fields(33) = 0
        rsGravaTMov.Fields(34) = 0
        rsGravaTMov.Fields(35) = 0
        rsGravaTMov.Fields(36) = Time
        rsGravaTMov.Fields(37) = 0
        rsGravaTMov.Fields(38) = 0
        rsGravaTMov.Fields(39) = 0
        rsGravaTMov.Fields(40) = 0
        rsGravaTMov.Fields(41) = 0
        rsGravaTMov.Fields(42) = vCodUsuarioRM
        rsGravaTMov.Fields(43) = DTPicker1.Value
        rsGravaTMov.Fields(44) = 0
        rsGravaTMov.Fields(45) = vValor
        rsGravaTMov.Fields(46) = 0
        rsGravaTMov.Fields(47) = 0
        rsGravaTMov.Fields(48) = 0
        rsGravaTMov.Fields(49) = 0
        rsGravaTMov.Fields(50) = vSequencialEstoque
        rsGravaTMov.Fields(51) = 0
        rsGravaTMov.Fields(52) = "T"
        rsGravaTMov.Fields(53) = DTPicker1.Value
        rsGravaTMov.Fields(54) = 0
        rsGravaTMov.Fields(55) = 0
        rsGravaTMov.Fields(56) = 539
        rsGravaTMov.Fields(57) = 0
        rsGravaTMov.Fields(58) = 0
        rsGravaTMov.Fields(59) = 0
        rsGravaTMov.Fields(60) = 0
        rsGravaTMov.Fields(61) = 0
        rsGravaTMov.Fields(62) = 0
        rsGravaTMov.Fields(63) = 0
        rsGravaTMov.Fields(64) = vCodUsuarioRM
        rsGravaTMov.Fields(65) = DTPicker1.Value
        rsGravaTMov.Fields(66) = vCodUsuarioRM
        rsGravaTMov.Fields(67) = DTPicker1.Value
    End If
 
    rsGravaTMov.Update
    rsGravaTMov.Close
    Set rsGravaTMov = Nothing
    cnBancoSAP.CommitTrans
    
    Exit Sub
Err:
    cnBancoSAP.RollbackTrans
End Sub

Private Sub GravaTitMMov(vSequencialItens As Integer)
On Error GoTo Err
    cnBancoSAP.BeginTrans
    
    Dim rsGravaTitMMov As New ADODB.Recordset
    Dim SqlGravaTitMMov As String
   
    SqlGravaTitMMov = "SELECT A.CODCOLIGADA,A.IDMOV,A.NSEQITMMOV,A.NUMEROSEQUENCIAL,A.IDPRD,A.QUANTIDADE,A.PRECOUNITARIO,A.PRECOTABELA,A.DATAEMISSAO,A.CODUND,A.QUANTIDADEARECEBER,A.FLAGEFEITOSALDO,A.VALORUNITARIO,A.VALORFINANCEIRO,A.ALIQORDENACAO,A.QUANTIDADEORIGINAL,A.FLAG,A.BLOCK,A.FATORCONVUND," & _
    "A.VALORTOTALITEM,A.CODFILIAL,A.QUANTIDADESEPARADA,A.PERCENTCOMISSAO,A.COMISSAOREPRES,A.VALORESCRITURACAO,A.VALORFINPEDIDO,A.VALOROPFRM1,A.VALOROPFRM2,A.PRECOEDITADO,A.QTDEVOLUMEUNITARIO,A.PRECOTOTALEDITADO,A.VALORDESCCONDICONALITM,A.VALORDESPCONDICIONALITM,A.VALORUNTORCAMENTO," & _
    "A.VALSERVICONFE,A.CODLOC,A.VALORBEM,A.VALORLIQUIDO,A.VALORBRUTOITEM,A.VALORBRUTOITEMORIG,A.QUANTIDADETOTAL,A.PRODUTOSUBSTITUTO,A.PRECOUNITARIOSELEC,A.RECCREATEDBY,A.RECCREATEDON,A.RECMODIFIEDBY,A.RECMODIFIEDON,A.QUANTIDADECONCLUIDA FROM TITMMOV AS A WHERE A.IDMOV = '" & vIDMov & "'"
    rsGravaTitMMov.Open SqlGravaTitMMov, cnBancoSAP, adOpenKeyset, adLockOptimistic
    
    'If rsGravaTitMMov.RecordCount = 0 Then
        rsGravaTitMMov.AddNew
        rsGravaTitMMov.Fields(0) = vCodColigadaRM ' Código da coligada RM
        rsGravaTitMMov.Fields(1) = vIDMov 'identificador do movimento RM
        rsGravaTitMMov.Fields(2) = vSequencialItens ' Sequencial dos itens
        rsGravaTitMMov.Fields(3) = vSequencialItens ' Sequencial do itens
        rsGravaTitMMov.Fields(4) = ListView2.SelectedItem.ListSubItems.Item(12) 'IDPRD
        rsGravaTitMMov.Fields(5) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(6) = ListView2.SelectedItem.ListSubItems.Item(11) / ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(7) = 0
        rsGravaTitMMov.Fields(8) = DTPicker1.Value
        rsGravaTitMMov.Fields(9) = ListView2.SelectedItem.ListSubItems.Item(4)
        rsGravaTitMMov.Fields(10) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(11) = 1
        rsGravaTitMMov.Fields(12) = ListView2.SelectedItem.ListSubItems.Item(11) / ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(13) = ListView2.SelectedItem.ListSubItems.Item(11) / ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(14) = 0
        rsGravaTitMMov.Fields(15) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(16) = 0
        rsGravaTitMMov.Fields(17) = 0
        rsGravaTitMMov.Fields(18) = 0
        rsGravaTitMMov.Fields(19) = 0
        rsGravaTitMMov.Fields(20) = 1
        rsGravaTitMMov.Fields(21) = 0
        rsGravaTitMMov.Fields(22) = 0
        rsGravaTitMMov.Fields(23) = 0
        rsGravaTitMMov.Fields(24) = 0
        rsGravaTitMMov.Fields(25) = 0
        rsGravaTitMMov.Fields(26) = 0
        rsGravaTitMMov.Fields(27) = 0
        rsGravaTitMMov.Fields(28) = 0
        rsGravaTitMMov.Fields(29) = 1
        rsGravaTitMMov.Fields(30) = 0
        rsGravaTitMMov.Fields(31) = 0
        rsGravaTitMMov.Fields(32) = 0
        rsGravaTitMMov.Fields(33) = 0
        rsGravaTitMMov.Fields(34) = 0
        rsGravaTitMMov.Fields(35) = vCodLocalEstoque 'local de estoque
        rsGravaTitMMov.Fields(36) = 0
        rsGravaTitMMov.Fields(37) = ListView2.SelectedItem.ListSubItems.Item(11) / ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(38) = ListView2.SelectedItem.ListSubItems.Item(11) / ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(39) = 0
        rsGravaTitMMov.Fields(40) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(41) = 0
        rsGravaTitMMov.Fields(42) = 0
        rsGravaTitMMov.Fields(43) = vCodUsuarioRM
        rsGravaTitMMov.Fields(44) = DTPicker1.Value
        rsGravaTitMMov.Fields(45) = vCodUsuarioRM
        rsGravaTitMMov.Fields(46) = DTPicker1.Value
        rsGravaTitMMov.Fields(47) = 0
    
        rsGravaTitMMov.Update
        rsGravaTitMMov.Close
        Set rsGravaTitMMov = Nothing
    'End If
    cnBancoSAP.CommitTrans
    Exit Sub
Err:
    cnBancoSAP.RollbackTrans
End Sub

Private Sub GravaTprdLoc()
    Dim rsGravaTprdLoc As New ADODB.Recordset
    Dim SqlGravaTprdLoc As String

    SqlGravaTprdLoc = "UPDATE TPRDLOC set SALDOFISICO2 = SALDOFISICO2+" & ListView2.SelectedItem.ListSubItems.Item(3) & " where codcoligada = '" & vCodColigadaRM & "' and CODLOC = '" & vCodLocalEstoque & "' AND CODFILIAL = 1 AND IDPRD = " & ListView2.SelectedItem.ListSubItems.Item(12)
    rsGravaTprdLoc.Open SqlGravaTprdLoc, cnBancoSAP

    Set rsGravaTprdLoc = Nothing

    SqlGravaTprdLoc = "UPDATE TPRDLOC set SALDOFINANCEIRO2 = SALDOFISICO2*CUSTOMEDIO where codcoligada = '" & vCodColigadaRM & "' and CODLOC = " & vCodLocalEstoque & " AND CODFILIAL = 1 AND IDPRD = '" & ListView2.SelectedItem.ListSubItems.Item(12) & "'"
    rsGravaTprdLoc.Open SqlGravaTprdLoc, cnBancoSAP
    
    Set rsGravaTprdLoc = Nothing

End Sub

Private Sub GravaRelacEntreMovCab()
On Error GoTo Err
    cnBancoSAP.BeginTrans
    
    Dim rsGravaRelacEntreMov As New ADODB.Recordset
    Dim SqlGravaRelacEntreMov As String
   
    SqlGravaRelacEntreMov = "select * from TMOVRELAC as A"
    rsGravaRelacEntreMov.Open SqlGravaRelacEntreMov, cnBancoSAP, adOpenKeyset, adLockOptimistic
    
    rsGravaRelacEntreMov.Fields(0) = vIDMov 'ID do movimento de devolução
    rsGravaRelacEntreMov.Fields(1) = vCodColigadaRM 'Código da coligada do movimento de devolução
    rsGravaRelacEntreMov.Fields(2) = vIDMovEmprestimo 'ID do movimento de empréstimo
    rsGravaRelacEntreMov.Fields(3) = vCodColigadaEmprestimo 'Código da coligada do movimento de empréstimo
    rsGravaRelacEntreMov.Fields(4) = "V" '? - Sempre vai ser V
    rsGravaRelacEntreMov.Fields(6) = vCodUsuarioRM ' Nome do usuário que criou o movimento
    rsGravaRelacEntreMov.Fields(7) = DTPicker1.Value 'Data que o movimento foi criado
    rsGravaRelacEntreMov.Fields(8) = vCodUsuarioRM  'Nome do usuário que alterou o movimento
    rsGravaRelacEntreMov.Fields(9) = DTPicker1.Value 'Data que o movimento foi alterado
    rsGravaRelacEntreMov.Update
    rsGravaRelacEntreMov.Close
    Set rsGravaRelacEntreMov = Nothing
    cnBancoSAP.CommitTrans
    Exit Sub
Err:
    cnBancoSAP.RollbackTrans
End Sub

Private Sub GravaRelacEntreMovItens(vIDMovEmp As Double, vIDPRD As Integer, vSequencialDev As Integer, vQuantDev As Double)
On Error GoTo Err
    cnBancoSAP.BeginTrans
    
    Dim rsGravaTitMMov As New ADODB.Recordset
    Dim SqlGravaTitMMov As String
    Dim vSequencialItem As Integer
    
    SqlGravaTitMMov = "SELECT * FROM TBEMPRESTIMOITENS AS A WHERE A.IDMOV = '" & vIDMovEmp & "' and A.IDPRD = '" & vIDPRD & "'"
    rsGravaTitMMov.Open SqlGravaTitMMov, cnBanco, adOpenKeyset, adLockReadOnly
    vSequencialItem = rsGravaTitMMov.Fields(13)
    
    
    Dim rsGravaRelacEntreMovItens As New ADODB.Recordset
    Dim SqlGravaRelacEntreMovItens As String

    SqlGravaRelacEntreMovItens = "select * from TITMMOVRELAC as A"
    rsGravaRelacEntreMovItens.Open SqlGravaRelacEntreMovItens, cnBancoSAP, adOpenKeyset, adLockOptimistic
    rsGravaRelacEntreMovItens.Fields(0) = vIDMov 'ID do movimento de devolução
    rsGravaRelacEntreMovItens.Fields(1) = vSequencialDev 'Sequencial do item de devolução
    rsGravaRelacEntreMovItens.Fields(2) = vCodColigadaRM 'Código da coligada do movimento de devolução
    rsGravaRelacEntreMovItens.Fields(3) = vIDMovEmprestimo 'ID do movimento de emprestimo
    rsGravaRelacEntreMovItens.Fields(4) = vSequencialItem 'Sequencial do item do emprestimo
    rsGravaRelacEntreMovItens.Fields(5) = vCodColigadaRM ' Código da coligada do movimento de emprestimo
    rsGravaRelacEntreMovItens.Fields(6) = vQuantDev 'Quantidade do movimento de devolução
    rsGravaRelacEntreMovItens.Fields(7) = vCodUsuarioRM ' Nome do usuário que criou o movimento
    rsGravaRelacEntreMovItens.Fields(8) = DTPicker1.Value 'Data que o movimento foi criado
    rsGravaRelacEntreMovItens.Fields(9) = vCodUsuarioRM 'Nome do usuário que alterou o movimento
    rsGravaRelacEntreMovItens.Fields(10) = DTPicker1.Value 'Data que o movimento foi alterado
    rsGravaRelacEntreMovItens.Update
    rsGravaRelacEntreMovItens.Close
    Set rsGravaRelacEntreMovItens = Nothing
    cnBancoSAP.CommitTrans
    Exit Sub
Err:
    cnBancoSAP.RollbackTrans
End Sub

