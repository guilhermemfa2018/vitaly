VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmprestimo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empréstimo"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmprestimo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   16080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "<F7> Devolução"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   40
      Top             =   8640
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Itens Empresatados (                            )"
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
      TabIndex        =   17
      Top             =   1200
      Width           =   15855
      Begin VB.Frame Frame7 
         Caption         =   "Histórico de Movimentações"
         Height          =   6975
         Left            =   8640
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   7215
         Begin MSComctlLib.ListView ListView4 
            Height          =   6615
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   11668
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   4080
         TabIndex        =   36
         Top             =   480
         Width           =   3735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "frmEmprestimo.frx":0CCA
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEmprestimo 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox chkEmprestimo 
         Caption         =   "Contém"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEmprestimo 
         Caption         =   "<"
         Height          =   615
         Index           =   2
         Left            =   7920
         TabIndex        =   9
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdEmprestimo 
         Caption         =   ">"
         Height          =   615
         Index           =   1
         Left            =   7920
         TabIndex        =   8
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   8640
         TabIndex        =   7
         Top             =   480
         Width           =   7095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   8640
         OleObjectBlob   =   "frmEmprestimo.frx":0D40
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtEmprestimo 
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Descrição ou código do produto"
         Top             =   480
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":0DBA
         TabIndex        =   18
         Top             =   240
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6255
         Left            =   8640
         TabIndex        =   10
         Top             =   960
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   11033
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11033
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6975
         Left            =   3120
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
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
      Left            =   13080
      TabIndex        =   38
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox txtEmprestimo 
      Height          =   375
      Index           =   9
      Left            =   4680
      TabIndex        =   37
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
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
      TabIndex        =   31
      Top             =   120
      Width           =   5655
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":0E50
         TabIndex        =   34
         Top             =   720
         Width           =   5415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":0F18
         TabIndex        =   33
         Top             =   480
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":0FBE
         TabIndex        =   32
         Top             =   240
         Width           =   3735
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
      TabIndex        =   13
      Top             =   120
      Width           =   10095
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   8400
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16121857
         CurrentDate     =   42632
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   8400
         OleObjectBlob   =   "frmEmprestimo.frx":1052
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdEmprestimo 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   3
         ToolTipText     =   "Realiza a pesquisa de todos os colaboradores"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtEmprestimo 
         Height          =   330
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Tag             =   "Nome do colaborador"
         ToolTipText     =   "Digite parte do NOME do colaborador e tecle ENTER"
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtEmprestimo 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Chapa do colaborador"
         ToolTipText     =   "Digite a CHAPA do colaborador"
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "frmEmprestimo.frx":10CA
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":112C
         TabIndex        =   14
         Top             =   240
         Width           =   1095
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
      Left            =   3480
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   2040
         TabIndex        =   26
         Top             =   480
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":1190
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmEmprestimo.frx":11F6
         TabIndex        =   29
         Top             =   240
         Width           =   1335
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
      Left            =   3480
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtEmprestimo 
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   2040
         TabIndex        =   21
         Top             =   480
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmEmprestimo.frx":1258
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmEmprestimo.frx":12C4
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin Ferramentaria.chameleonButton cmdEmp 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   12
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
      MICON           =   "frmEmprestimo.frx":132C
      PICN            =   "frmEmprestimo.frx":1348
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Ferramentaria.chameleonButton cmdEmp 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   11
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
      MICON           =   "frmEmprestimo.frx":2022
      PICN            =   "frmEmprestimo.frx":203E
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
Attribute VB_Name = "frmEmprestimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPosicaoAtual As Integer
Private rsCriterio As New ADODB.Recordset
Private SqlCriterio As String

Private Sub cmdEmp_Click(Index As Integer)
    Select Case Index
    Case 0
        'Msgbox "Rotina em desenvolvimento", vbInformation, "Ferramentaria"
        'Exit Sub
        
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "Ferramentaria"
            varGlobal = txtEmprestimo(0).Text
            FCREmprestimo.Show 1
            Unload Me
        Else
            SkinLabel1.Visible = False
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "Ferramentaria"
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub cmdEmprestimo_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridChapa ("")
        chamaChapa "" 'txtEmprestimo(1).Text
    Case 1
        addRemLoteNota ListView1, ListView2
        vQtdSolicitada = 0
        SomaLV ListView2, 6, Text1
    Case 2
        addRemLoteNota ListView2, ListView1
        SomaLV ListView2, 6, Text1
    Case 3
        PesquisaProd
        ApagaExceso
    End Select
End Sub

Private Sub PesquisaProd()
    'Dim vLocalEstoque As String
    LimpaLV ListView1
    Dim vTxtPesquisa As String, vPesqCampo As String
    If IsNumeric(txtEmprestimo(6).Text) Then
        vPesqCampo = "Codigo"
    Else
        vPesqCampo = "Descricao"
    End If
    
    If chkEmprestimo.Value = 0 Then
        vTxtPesquisa = txtEmprestimo(6).Text & "%"
    Else
        vTxtPesquisa = "%" & txtEmprestimo(6).Text & "%"
    End If
    
    If vPesqCampo = "Descricao" Then
    chamaSQL "select C.CODLOC,A.CODIGOPRD,(C.SALDOFISICO2-ISNULL(C.SALDOFISICO6,0)) AS SALDOFISICO2,B.CODUNDCONTROLE,A.NOMEFANTASIA,case when  D.DATAVENCIMENTO <= getdate() then 'Sim' else 'Não'  end Manut_Venc,max(c.CUSTOMEDIO),A.IDPRD from " & vBancoSAP & ".DBO.TPRODUTO AS A inner join " & vBancoSAP & ".DBO.TPRODUTODEF AS B on B.IDPRD=A.IDPRD inner join " & vBancoSAP & ".DBO.TPRDLOC AS C on C.IDPRD=A.IDPRD " & _
             "left join " & vBancoSAP & ".DBO.OFVENCPLANOMANUT AS D on D.IDOBJOF = SUBSTRING(A.CODIGOPRD,4,9) left join " & vBancoSAP & ".DBO.OFVENCPLANOMANUT AS E on E.IDOBJOF = D.IDOBJOF left join " & vBancoSAP & ".DBO.OFPLANOMANUT AS F on F.IDPLANO = E.idplano and F.ATIVO = 1 where C.CODCOLIGADA=1 and A.INATIVO=0 and TIPO='P' and (A.CODIGOPRD like '01.%' or A.OBSERVACAO='FERRAMENTA' or A.CODIGOPRD like '04.%' OR A.CODIGOPRD like '03.0001.1752' OR A.CODIGOPRD " & _
             "like '03.0001.1722' OR A.CODIGOPRD like '03.0001.1682' OR A.CODIGOPRD like '03.0001.3112' OR A.CODIGOPRD like '03.0001.3114') and C.SALDOFISICO2-ISNULL(C.SALDOFISICO6,0)>0 and A.NOMEFANTASIA like '" & vTxtPesquisa & "' and C.CODLOC='" & vLocalEstoque & "' group by A.CODCOLPRD,A.NOMEFANTASIA,A.CODIGOPRD,C.SALDOFISICO2,C.SALDOFISICO6,B.CODUNDCONTROLE,A.IDPRD,B.PRECO1,C.CODLOC,D.DATAVENCIMENTO order by A.NOMEFANTASIA"
    Else
    chamaSQL "select C.CODLOC,A.CODIGOPRD,(C.SALDOFISICO2-ISNULL(C.SALDOFISICO6,0)) AS SALDOFISICO2,B.CODUNDCONTROLE,A.NOMEFANTASIA,case when  D.DATAVENCIMENTO <= getdate() then 'Sim' else 'Não'  end Manut_Venc,max(c.CUSTOMEDIO),A.IDPRD from " & vBancoSAP & ".DBO.TPRODUTO AS A inner join " & vBancoSAP & ".DBO.TPRODUTODEF AS B on B.IDPRD=A.IDPRD inner join " & vBancoSAP & ".DBO.TPRDLOC AS C on C.IDPRD=A.IDPRD " & _
             "left join " & vBancoSAP & ".DBO.OFVENCPLANOMANUT AS D on D.IDOBJOF = SUBSTRING(A.CODIGOPRD,4,9) left join " & vBancoSAP & ".DBO.OFVENCPLANOMANUT AS E on E.IDOBJOF = D.IDOBJOF left join " & vBancoSAP & ".DBO.OFPLANOMANUT AS F on F.IDPLANO = E.idplano and F.ATIVO = 1 where C.CODCOLIGADA=1 and A.INATIVO=0 and TIPO='P' and (A.CODIGOPRD like '01.%' or A.OBSERVACAO='FERRAMENTA' or A.CODIGOPRD like '04.%' OR A.CODIGOPRD like '03.0001.1752' OR A.CODIGOPRD " & _
             "like '03.0001.1722' OR A.CODIGOPRD like '03.0001.1682' OR A.CODIGOPRD like '03.0001.3112' OR A.CODIGOPRD like '03.0001.3114') and C.SALDOFISICO2-ISNULL(C.SALDOFISICO6,0)>0 and A.CODIGOPRD like '" & vTxtPesquisa & "' and C.CODLOC='" & vLocalEstoque & "' group by A.CODCOLPRD,A.NOMEFANTASIA,A.CODIGOPRD,C.SALDOFISICO2,C.SALDOFISICO6,B.CODUNDCONTROLE,A.IDPRD,B.PRECO1,C.CODLOC,D.DATAVENCIMENTO order by A.NOMEFANTASIA"
    End If
    Compoe_Listview ListView1, Sqlp, "0000"
    

    
    PersonaColLVForm ListView1, 6, "N", "N", "", "N", "N", "S", "D"
    PersonaColLVForm ListView1, 5, "N", "S", "", "N", "N", "N", "D"
End Sub

Private Sub Command1_Click()
    If txtEmprestimo(0).Text = "" Or txtEmprestimo(1).Text = "" Then
        mobjMsg.Abrir "os campos CHAPA e NOME do colaborador devem estar preenchidos", Ok, informacao, "Ferramentaria"
    Else
        frmDevolucao.Show 1
    End If
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
            FCREmprestimo.Show 1
            Unload Me
        'Else
            'SkinLabel1.Visible = False
            'mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "Ferramentaria"
        End If
    ElseIf KeyCode = 118 Then
        If txtEmprestimo(0).Text = "" Or txtEmprestimo(1).Text = "" Then
            mobjMsg.Abrir "os campos CHAPA e NOME do colaborador devem estar preenchidos", Ok, informacao, "Ferramentaria"
        Else
            frmDevolucao.Show 1
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Me.KeyPreview = True
    Status = Pesquisa
    listview_cabecalho
    If varGlobal <> 0 And Len(varGlobal) <> 12 Then
        txtEmprestimo(0).Text = varGlobal
        If chamaChapa("") = False Then Exit Sub
    End If
    If Status = "novo" Then
        DTPicker1.Value = Date
        txtEmprestimo(7).Text = Mid$(Principal.Caption, 42, 50)
        txtEmprestimo(8).Text = vCodVenRM & " - " & vNomeVenRM 'NomUsu
        vCodLocalEstoque = vLocalEstoque
    ElseIf Status = "editar" Then
        Me.Caption = "Empréstimo - Visualização do movimento: " & Mid$(varGlobal, 7, 6)
        txtEmprestimo(7).Text = Mid$(Principal.Caption, 42, 50)
        vCodLocalEstoque = vLocalEstoque
        ResultPesq
        chamaSQL "Select a.localestoque,a.codigoprd,a.qtdemprestado,a.um,a.descricao,'NÃO',a.valortotal,a.idprd from tbEmprestimoItens as a inner join tbEmprestimo as b on a.idmov = b.idmov and a.numeromov = b.numeromov where a.codcoligada = 1 and a.localestoque = " & vLocalEstoque & " and a.chapa = '" & Mid$(varGlobal, 1, 6) & "' and b.numeromov = '" & Mid$(varGlobal, 7, 6) & "'"
        Compoe_Listview ListView2, Sqlp, "00"
        PersonaColLVForm ListView2, 6, "N", "N", "", "N", "N", "S", "D"
        BloqControl
        SomaLV ListView2, 6, Text1
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

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "LOC", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "COD PRD", ListView1.Width / 5.5
    ListView1.ColumnHeaders.Add , , "QTDE", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "UN", ListView1.Width / 18
    ListView1.ColumnHeaders.Add , , "DESCRIÇÃO", ListView1.Width / 2.5
    ListView1.ColumnHeaders.Add , , "M. VENC", ListView1.Width / 9
    ListView1.ColumnHeaders.Add , , "VALOR", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "IDPRD", ListView1.Width / 10000
    Me.ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
'    Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
'    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "LOC", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "COD PRD", ListView2.Width / 5
    ListView2.ColumnHeaders.Add , , "QTDE", ListView2.Width / 11.5
    ListView2.ColumnHeaders.Add , , "UN", ListView2.Width / 17
    ListView2.ColumnHeaders.Add , , "DESCRIÇÃO", ListView2.Width / 2.5
    ListView2.ColumnHeaders.Add , , "M. VENC", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "VALOR", ListView2.Width / 7
    ListView2.ColumnHeaders.Add , , "IDPRD", ListView2.Width / 10000
    Me.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

    'Listview de Devolucao. Serve para o sistema procurar por itens que precisam ser devolvidos, pois, precisam ir para manutencao
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "LOC", ListView3.Width / 15
    ListView3.ColumnHeaders.Add , , "ID MOV", ListView3.Width / 14
    ListView3.ColumnHeaders.Add , , "COD PRD", ListView3.Width / 8.8
    ListView3.ColumnHeaders.Add , , "QTDE", ListView3.Width / 17.5
    ListView3.ColumnHeaders.Add , , "UN", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "DESCRIÇÃO", ListView3.Width / 4
    ListView3.ColumnHeaders.Add , , "DATA EMP", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "DIAS EMP", ListView3.Width / 12
    ListView3.ColumnHeaders.Add , , "DIAS OBRA", ListView3.Width / 10.5
    ListView3.ColumnHeaders.Add , , "MNT P", ListView3.Width / 15.5
    ListView3.ColumnHeaders.Add , , "RECOLHE", ListView3.Width / 13
    ListView3.ColumnHeaders.Add , , "VALOR", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "ID PRD", ListView3.Width / 10000
    Me.ListView3.ColumnHeaders(4).Alignment = lvwColumnRight
    
    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "MOV", ListView4.Width / 15
    ListView4.ColumnHeaders.Add , , "CHAPA", ListView4.Width / 14
    ListView4.ColumnHeaders.Add , , "FUNCIONARIO", ListView4.Width / 8.8
    ListView4.ColumnHeaders.Add , , "DATA", ListView4.Width / 17.5
    ListView4.ColumnHeaders.Add , , "USUARIO", ListView4.Width / 10000
    
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    
    ListView4.BackColor = RGB(135, 194, 194)
    
End Sub

Private Sub ListView1_Click()
    MarcaDesmarca ListView1
End Sub

Private Sub ListView1_DblClick()
    addRemLoteNota ListView1, ListView2
    vQtdSolicitada = 0
    DoEvents
    SomaLV ListView2, 6, Text1
    ListView1.SetFocus
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    MarcaDesmarca ListView1
    If KeyCode = 13 Or KeyCode = 9 Or KeyCode = 32 Then ' Enter ou TAB
        addRemLoteNota ListView1, ListView2
        vQtdSolicitada = 0
        DoEvents
        SomaLV ListView2, 6, Text1
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
        SomaLV ListView2, 6, Text1
        ListView2.SetFocus
    End If
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
    MarcaDesmarca ListView2
End Sub

Private Sub txtEmprestimo_GotFocus(Index As Integer)
    mudaCorText txtEmprestimo(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtEmprestimo.Count - 1
        txtEmprestimo(X).SelStart = 0
        txtEmprestimo(X).SelLength = Len(txtEmprestimo(X).Text)
    Next
End Sub

Private Sub txtEmprestimo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtEmprestimo(0).Text = Format(txtEmprestimo(0).Text, "000000")
            If chamaChapa("") = False Then Exit Sub
        End If
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            chamaChapa txtEmprestimo(1).Text
        End If
    Case 6
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            PesquisaProd
            ApagaExceso
        End If
    End Select
End Sub

Private Function chamaChapa(vNome As String)
On Error GoTo Err
    chamaChapa = False
    Dim rschamaChapa As New ADODB.Recordset
    Dim SqlchamaChapa As String
    
    If vNome = "" Then
        SqlchamaChapa = "select a.CODCOLIGADA,a.CODVEN,a.NOME,a.INATIVO,b.SITEMPRESTIMO,b.MOTBLOQUEIO,c.CODFUNCAO,d.NOME,c.CODSECAO,e.DESCRICAO from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b  on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO where a.CODCOLIGADA=1 and a.INATIVO=0 and a.codven = " & txtEmprestimo(0).Text & " AND C.CODSITUACAO in('A','F','P','Z') order by a.nome"
    Else
        SqlchamaChapa = "select a.CODCOLIGADA,a.CODVEN,a.NOME,a.INATIVO,b.SITEMPRESTIMO,b.MOTBLOQUEIO,c.CODFUNCAO,d.NOME,c.CODSECAO,e.DESCRICAO from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b  on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO where a.CODCOLIGADA=1 and a.INATIVO=0 and a.nome like '" & vNome & "%' AND C.CODSITUACAO in('A','F','P','Z') order by a.nome"
        ChamaGridChapa (SqlchamaChapa)
        SqlchamaChapa = "select a.CODCOLIGADA,a.CODVEN,a.NOME,a.INATIVO,b.SITEMPRESTIMO,b.MOTBLOQUEIO,c.CODFUNCAO,d.NOME,c.CODSECAO,e.DESCRICAO from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b  on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA left join " & vBancoSAP & ".dbo.PFUNCAO as d on c.CODFUNCAO = d.CODIGO left join " & vBancoSAP & ".dbo.PSECAO as e on c.CODSECAO = e.CODIGO where a.CODCOLIGADA=1 and a.INATIVO=0 and a.codven =" & Pesquisa & " AND C.CODSITUACAO in('A','F','P','Z') order by a.nome"
        vNome = ""
        'Exit Function
    End If
    rschamaChapa.Open SqlchamaChapa, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rschamaChapa.EOF Then
        txtEmprestimo(0).Text = Format(txtEmprestimo(0).Text, "000000")
        txtEmprestimo(1).Text = rschamaChapa.Fields(2)  'Nome
        txtEmprestimo(2).Text = rschamaChapa.Fields(6)  'cod função
        txtEmprestimo(3).Text = rschamaChapa.Fields(7)  'Nome função
        txtEmprestimo(4).Text = rschamaChapa.Fields(8)  'cod setor
        txtEmprestimo(5).Text = rschamaChapa.Fields(9)  'Nome setor
        varGlobal = txtEmprestimo(0).Text
        CompoeControles = True
    Else
        mobjMsg.Abrir "Registro de colaborador não identificado no sistema", Ok, critico, "Atenção"
        txtEmprestimo(0).Text = ""
        txtEmprestimo(1).Text = ""
        txtEmprestimo(0).SetFocus
    End If
    
'----------------------------------------
'Area destinada a compor o listview de Devolucao
    ListView3.ListItems.Clear
    If txtEmprestimo(0).Text <> "" Then
        chamaSQL "select b.localestoque as codloc,a.idmov,b.codigoprd,(b.qtdemprestado-b.qtddevolvida) as QTDEPENDENTE,b.um,b.descricao,a.dataemprestimo as DATAEMISSAO,qtDiasEmp = p.CAMPOLIVRE ,dife = CONVERT(VARCHAR, DATEDIFF(DAY, M.DATAEMISSAO ,GETDATE()) ) ,manutencao = ( case when ( SELECT manu.DATAVENCIMENTO from " & vBancoSAP & ".dbo.OFVENCPLANOMANUT manu INNER join " & vBancoSAP & ".dbo.TPRODUTO Prd on manu.IDOBJOF = SUBSTRING(Prd.CODIGOPRD,4,9) AND PRD.CODIGOPRD  = P.CODIGOPRD) < GETDATE()then 'Sim' else 'Não' end),recolher  =  case when (p.CAMPOLIVRE - CONVERT(VARCHAR, DATEDIFF(DAY, a.dataemprestimo ,GETDATE()) )) <0 then  'Sim' else 'Não' end, " & _
                 "b.valortotal/b.qtdemprestado as valor_unit,b.idprd from tbEmprestimo as a inner join tbEmprestimoItens as b on a.idmov = b.idmov and (b.qtdemprestado-b.qtddevolvida) > 0 inner join " & vBancoSAP & ".dbo.tloc as c on b.localestoque = c.CODLOC COLLATE SQL_Latin1_General_CP1_CI_AS and c.CODFILIAL = 1 inner join " & vBancoSAP & ".dbo.TMOV as m on a.codcoligada = m.CODCOLIGADA and CAST(a.numeromov AS INT) = m.NUMEROMOV COLLATE SQL_Latin1_General_CP1_CI_AS and a.serie = m.SERIE COLLATE SQL_Latin1_General_CP1_CI_AS and a.idmov = m.IDMOV and m.CODFILIAL = 1 inner join " & vBancoSAP & ".dbo.TPRODUTO P on b.idprd = p.IDPRD where a.codcoligada = 1 and a.chapa = '" & txtEmprestimo(0).Text & "' AND A.localestoque = " & Val(vLocalEstoque) & ""
        Compoe_Listview ListView3, Sqlp, "00"
        PersonaColLVForm ListView3, 10, "N", "N", "", "N", "N", "S", "D"
        PersonaColLVForm ListView3, 11, "N", "S", "", "N", "N", "N", "D"
        
        Dim YDev As Integer, XDev As Integer
        YDev = ListView3.ListItems.Count
        'VERIFICA SE EXISTEM PRODUTOS VENCIDOS DE DEVOLUCAO
        For XDev = 1 To YDev
            If YDev < XDev Then
                Exit For
            End If
            ListView3.ListItems.Item(XDev).Selected = True 'Passar a selecao para o próximo item
            If ListView3.SelectedItem.ListSubItems.Item(10) = "Sim" Then
                If GeraLog = "S" Then
                    mobjMsg.Abrir "O Colaborador precisa devolver produto(s) vencido(s)", Ok, critico, "BLOQUEADO"
                    chamaChapa = False
                    'Envia para a tela de Devolução
                    frmDevolucao.Show 1
                    'Limpa os controles
                    txtEmprestimo(0).Text = ""
                    txtEmprestimo(1).Text = ""
                    txtEmprestimo(0).SetFocus
                    Exit Function
                Else
                    mobjMsg.Abrir "O Colaborador possui produto(s) vencido(s)", Ok, informacao, "ATENÇÃO"
                    chamaChapa = True
                    txtEmprestimo(6).SetFocus
                    Exit Function
                End If
            End If
        Next
    End If
'----------------------------------------
    rschamaChapa.Close
    Set rschamaChapa = Nothing
Err:
    Exit Function
End Function

Private Sub ChamaGridChapa(vSqlp As String)
On Error GoTo Err
    Dim F As New frmPesqger2
    If vSqlp = "" Then
        Sqlp = "select a.CODCOLIGADA,a.CODVEN,a.NOME,a.INATIVO,b.SITEMPRESTIMO,b.MOTBLOQUEIO from " & vBancoSAP & ".dbo.TVEN as a left join " & vBancoSAP & ".dbo.TVENCOMPL as b  on b.CODCOLIGADA=a.CODCOLIGADA and b.CODVEN=a.CODVEN left join " & vBancoSAP & ".dbo.PFUNC as c on a.CODVEN = c.CHAPA where a.CODCOLIGADA=1 and a.INATIVO=0 AND C.CODSITUACAO in('A','F','P','Z')"
    Else
        Sqlp = vSqlp
        vSqlp = ""
    End If
    procnom = "nome"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Funcionários"
    Pesquisa = frmEmprestimo.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "codven=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If Pesquisa = "Pesquisa de Funcionários" Then Pesquisa = ""
            txtEmprestimo(0) = Format(Pesquisa, "000000")
            txtEmprestimo(1) = rsLocal.Fields(2)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
Err:
    If Err.Number = 3705 Then
        rsLocal.Close
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        Resume Next
    End If
End Sub

Private Sub txtEmprestimo_LostFocus(Index As Integer)
    voltaCorText txtEmprestimo(Index)
    Select Case Index
    Case 0
        txtEmprestimo(0).Text = Format(txtEmprestimo(0).Text, "000000")
        If chamaChapa("") = False Then Exit Sub
    End Select
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
                If lvOrigem.SelectedItem.ListSubItems.Item(5) = "Sim" Then
                    'mobjMsg.Abrir "Item com manuteção vencida.Empréstimo não autorizado", ok, informacao, "Atenção"
                    Msgbox "Item com manuteção vencida.Empréstimo não autorizado", vbInformation, "Atenção"
                    Exit Sub
                End If
                If lvOrigem.SelectedItem.ListSubItems.Item(2) > 1 Then
                    vQtdDisponivel = lvOrigem.SelectedItem.ListSubItems.Item(2)
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
                    If lvOrigem.SelectedItem.ListSubItems.Item(1) = lvDestino.SelectedItem.ListSubItems.Item(1) Then
                        lvDestino.SelectedItem.ListSubItems.Item(2) = Val(lvDestino.SelectedItem.ListSubItems.Item(2)) + vQtdSolicitada
                        lvDestino.SelectedItem.ListSubItems.Item(6) = Format(lvDestino.SelectedItem.ListSubItems.Item(6) * lvDestino.SelectedItem.ListSubItems.Item(2), "#,##0.00;(#,##0.00)") 'Valor total
                        If vQtdSolicitada = Val(lvOrigem.SelectedItem.ListSubItems.Item(2)) Then
                            lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                        Else
                            lvOrigem.SelectedItem.ListSubItems.Item(2) = Val(lvOrigem.SelectedItem.ListSubItems.Item(2)) - vQtdSolicitada
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
                    If lvOrigem.SelectedItem.ListSubItems.Item(1) = lvDestino.SelectedItem.ListSubItems.Item(1) Then
                        lvDestino.SelectedItem.ListSubItems.Item(2) = Val(lvDestino.SelectedItem.ListSubItems.Item(2)) + Val(lvOrigem.SelectedItem.ListSubItems.Item(2))
                        lvDestino.SelectedItem.ListSubItems.Item(6) = Format(lvDestino.SelectedItem.ListSubItems.Item(6) * lvDestino.SelectedItem.ListSubItems.Item(2), "#,##0.00;(#,##0.00)") 'Valor total
                        
                        lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                        Exit Sub
                    End If
                Next
            End If
            
            Set ItemLst = lvDestino.ListItems.Add(, , lvOrigem.ListItems(X)) ' Local de estoque
            ItemLst.SubItems(1) = "" & lvOrigem.SelectedItem.ListSubItems.Item(1) 'Código do produto
            If lvOrigem.Name = "ListView1" Then
                If lvOrigem.SelectedItem.ListSubItems.Item(2) > 1 Then
                    ItemLst.SubItems(2) = "" & vQtdSolicitada 'Quantidade
                Else
                    ItemLst.SubItems(2) = "" & lvOrigem.SelectedItem.ListSubItems.Item(2) 'Quantidade
                End If
            Else
                ItemLst.SubItems(2) = "" & lvOrigem.SelectedItem.ListSubItems.Item(2) 'Quantidade
            End If
            ItemLst.SubItems(3) = "" & lvOrigem.SelectedItem.ListSubItems.Item(3) 'Unidade de medida
            ItemLst.SubItems(4) = "" & lvOrigem.SelectedItem.ListSubItems.Item(4) 'Descrição do produto
            ItemLst.SubItems(5) = "" & lvOrigem.SelectedItem.ListSubItems.Item(5) 'Manutenção Vencida?
            ItemLst.SubItems(6) = "" & Format(lvOrigem.SelectedItem.ListSubItems.Item(6) * ItemLst.SubItems(2), "#,##0.00;(#,##0.00)") 'Valor total
            ItemLst.SubItems(7) = "" & lvOrigem.SelectedItem.ListSubItems.Item(7) 'Identificador do produto
            If lvOrigem.Name = "ListView1" Then
                If vQtdDisponivel = vQtdSolicitada Then
                    lvOrigem.ListItems.Remove (X) ' Remove item selecionado do ListView1
                Else
                    lvOrigem.SelectedItem.ListSubItems.Item(2) = lvOrigem.SelectedItem.ListSubItems.Item(2) - vQtdSolicitada
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

Private Sub ApagaExceso()
    On Error GoTo TrataErro
    Dim X As Integer, Y As Integer, X1 As Integer, Y1 As Integer
    Dim vCodPRD As String
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        vCodPRD = ListView2.SelectedItem.ListSubItems.Item(1)
        Y1 = ListView1.ListItems.Count
        If Y1 = 0 Then Exit Sub
        For X1 = 1 To Y1
            ListView1.ListItems.Item(X1).Selected = True 'Passar a selecao para o próximo item
            If ListView1.SelectedItem.ListSubItems.Item(1) = vCodPRD Then
                If ListView1.SelectedItem.ListSubItems.Item(2) = ListView2.SelectedItem.ListSubItems.Item(2) Then
                    ListView1.ListItems.Remove (X1)
                Else
                    ListView1.SelectedItem.ListSubItems.Item(2) = Val(ListView1.SelectedItem.ListSubItems.Item(2)) - Val(ListView2.SelectedItem.ListSubItems.Item(2))
                End If
            End If
        Next
    Next
TrataErro:
    If ListView1.ListItems.Count > 0 Then ListView1.ListItems.Item(1).Selected = True
    Resume Next
End Sub

Private Function salvar_Dados()
'On Error GoTo Err
    If ValidaCampo = False Then Exit Function
    vTransacaoAtiva = 1
    cnBanco.BeginTrans
    
    salvar_Dados = True
    
    GeraNumeroMov
    GeraIDMov
    GeraSequencialEstoque
    
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    limpaQualquerDado
    vQualquerDado(1, 1) = txtEmprestimo(0).Text 'Identificador do colaborador que pegou a ferramenta emprestada
    vQualquerDado(1, 2) = "S"
    vQualquerDado(2, 1) = txtEmprestimo(1) 'Nome do colaborador que pegou a ferramenta emprestada
    vQualquerDado(2, 2) = "S"
    vQualquerDado(3, 1) = txtEmprestimo(2).Text ' Identificador da função do colaborador que pegou a ferramenta emprestada
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtEmprestimo(3).Text ' Nome da função do colaborador que pegou a ferramenta emprestada
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtEmprestimo(4).Text ' Identificador da setor do colaborador que pegou a ferramenta emprestada
    vQualquerDado(5, 2) = "S"
    vQualquerDado(6, 1) = txtEmprestimo(5).Text ' Nome da setor do colaborador que pegou a ferramenta emprestada
    vQualquerDado(6, 2) = "S"
    If DTPicker1.Value <> "" Then
        vQualquerDado(7, 1) = DTPicker1.Value 'Data de realização do emprestimo para o colaborador
        vQualquerDado(7, 2) = "D"
    End If
    
    vQualquerDado(8, 1) = vIDMov ' VIDMOV - Identificador do movimento gerado pelo sistema deferramentaria
    vQualquerDado(8, 2) = "I"
    
    
    vQualquerDado(9, 1) = Format(vNumeromov, "000000") ' VNUMEROMOV - Numero do movimento gerado pelo sistema deferramentaria
    vQualquerDado(9, 2) = "S"
    
    vQualquerDado(10, 1) = vSerie ' SERIE - Serie do movimento da ferramentaria
    vQualquerDado(10, 2) = "S"
    
    
    vQualquerDado(11, 1) = "E" ' Status do critério
    vQualquerDado(11, 2) = "S"
    
'    If Check1.Value = 1 Then
'        vQualquerDado(11, 1) = "S" ' Status do critério
'    Else
'        vQualquerDado(11, 1) = "N" ' Status do critério
'    End If
'    vQualquerDado(11, 2) = "S"
    
    vQualquerDado(12, 1) = 1 ' Código da Coligada
    vQualquerDado(12, 2) = "I"
    
    vQualquerDado(13, 1) = vLocalEstoque ' Local de estoque
    vQualquerDado(13, 2) = "I"
    
    vQualquerDado(14, 1) = txtEmprestimo(8).Text ' Nome de quem emprestou
    vQualquerDado(14, 2) = "S"
    
    vQualquerDado(15, 1) = vCodUsuarioRM ' codusuario (RM) de quem emprestou
    vQualquerDado(15, 2) = "S"
    
    txtEmprestimo(9) = vNumeromov
    GravaDados "tbEmprestimo", "numeromov", "S", txtEmprestimo(9), 15, "", "", txtEmprestimo(9)
        
        
        
    'Grava dados ListView1
    
'    InsereCaracter "S", "N", ""
    limpaQualquerDado
'    ordenaLVArray ListView2, "4", "5", "0", "1", "2", "3", "6", "9", "", "", "", "", "", "", "", ""
    GravaProdutosEmprestimo

'    GravaDadosLV "tb", "idfornecedor", "S", txtFornecedor(0)
    
'    InsereCaracter "", "", ""
'    AtualizaListview
    cnBanco.CommitTrans
    vTransacaoAtiva = 0
    Exit Function
Err:
    salvar_Dados = False
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If txtEmprestimo(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtEmprestimo(0).Tag, Ok, critico, "Atenção"
        Me.txtEmprestimo(1).SetFocus
        Exit Function
    End If
    If txtEmprestimo(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtEmprestimo(1).Tag, Ok, critico, "Atenção"
        Me.txtEmprestimo(1).SetFocus
        Exit Function
    End If
    
    If ListView2.ListItems.Count = 0 Then
        mobjMsg.Abrir "Nenhuma ferramenta foi emprestada", Ok, critico, "Atenção"
        Me.txtEmprestimo(6).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function GeraNumeroMov()
    Dim rsGeraNumeroMov As New ADODB.Recordset
    Dim SqlGeraNumeroMov As String
    'vLocalEstoque
    SqlGeraNumeroMov = "Select top 1 * from tbMov as a where a.serie = 'FERE'+ '" & vLocalEstoque & "'  and codcoligada = 1 order by numeromov Desc"
    rsGeraNumeroMov.Open SqlGeraNumeroMov, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraNumeroMov.RecordCount > 0 Then
        vNumeromov = Val(rsGeraNumeroMov.Fields(0)) + 1
    Else
        vNumeromov = 1
    End If
    vSerie = "FERE" & vLocalEstoque
    rsGeraNumeroMov.Close
    Set rsGeraNumeroMov = Nothing

    limpaQualquerDado
    vQualquerDado(1, 1) = Format(vNumeromov, "000000") ' VNUMEROMOV - Numero do movimento gerado pelo sistema deferramentaria
    vQualquerDado(1, 2) = "S"
    vQualquerDado(2, 1) = vSerie ' SERIE - Serie do movimento da ferramentaria
    vQualquerDado(2, 2) = "S"
    vQualquerDado(3, 1) = 1 ' Código da Coligada
    vQualquerDado(3, 2) = "I"
    'GRAVA DADOS DO EMPRESTIMO
    GravaDados "tbMov", "Numeromov", "S", txtEmprestimo(0), 3, "", "", txtEmprestimo(0)
    
    'GRAVA DADOS DOS ITENS DO EMPRESTIMO
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
    rsAtualizaIDMov.Open SqlAtualizaIDMov, cnBancoSAP
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


Private Sub GravaProdutosEmprestimo()
    Dim rsGravaProdutosEmprestimo As New ADODB.Recordset
    Dim sqlGravaProdutosEmprestimo As String
    Dim X As Integer, Y As Integer
    
    sqlGravaProdutosEmprestimo = "Select * from tbEmprestimoItens as a where a.chapa = '" & txtEmprestimo(0).Text & "' and a.codcoligada = 1 and a.numeromov = '" & vNumeromov & "'"
    rsGravaProdutosEmprestimo.Open sqlGravaProdutosEmprestimo, cnBanco, adOpenKeyset, adLockOptimistic
    
    ListView2.ListItems.Item(1).Selected = True
    
    GravaTMov 'grava dados na tabela TMOV (TOTVS RM)
    'Exit Sub
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaProdutosEmprestimo.AddNew
        rsGravaProdutosEmprestimo(0) = txtEmprestimo(0).Text 'Chapa colaborador
        rsGravaProdutosEmprestimo(1) = ListView2.ListItems.Item(X) 'Local de estoque
        rsGravaProdutosEmprestimo(2) = ListView2.SelectedItem.ListSubItems.Item(1) 'Código do produto
        rsGravaProdutosEmprestimo(3) = ListView2.SelectedItem.ListSubItems.Item(4) 'Descrição do produto
        rsGravaProdutosEmprestimo(4) = vIDMov 'Identificador do movimento
        rsGravaProdutosEmprestimo(5) = ListView2.SelectedItem.ListSubItems.Item(7) 'Identificador do produto
        rsGravaProdutosEmprestimo(6) = ListView2.SelectedItem.ListSubItems.Item(2) 'quantidade emprestado
        
        rsGravaProdutosEmprestimo(7) = 0 'quantidade devolvida
        rsGravaProdutosEmprestimo(8) = ListView2.SelectedItem.ListSubItems.Item(2) 'quantidade pendente
        rsGravaProdutosEmprestimo(9) = DTPicker1.Value 'Data do empréstimo
        rsGravaProdutosEmprestimo(10) = Time 'Hora do empréstimo
        rsGravaProdutosEmprestimo(11) = "E" 'Status
        rsGravaProdutosEmprestimo(12) = NomUsu 'Nome de quem emprestou as ferramentas
        rsGravaProdutosEmprestimo(13) = X 'Sequencial
        rsGravaProdutosEmprestimo(14) = 1 'coligada
        rsGravaProdutosEmprestimo(15) = ListView2.SelectedItem.ListSubItems.Item(3) 'Unidade de medida do produto
        rsGravaProdutosEmprestimo(16) = ListView2.SelectedItem.ListSubItems.Item(6) 'Valor total
        rsGravaProdutosEmprestimo(17) = Format(vNumeromov, "000000") 'Numero do Movimento
        rsGravaProdutosEmprestimo(18) = vSerie 'Serie do movimento
        GravaTitMMov X 'grava dados na tabela TITMMOV (TOTVS RM)
        GravaTprdLoc 'grava dados na tabela TPRDLOC (TOTVS RM)
    Next
    If Not rsGravaProdutosEmprestimo.EOF Then rsGravaProdutosEmprestimo.Update
    rsGravaProdutosEmprestimo.Close
End Sub

Private Sub ResultPesq()
    SqlCriterio = "Select a.chapa,a.nome,dataemprestimo,a.numeromov,a.nomequememprestou from tbEmprestimo as a where codcoligada = 1 and a.localestoque ='" & Val(vLocalEstoque) & "' and a.chapa = '" & Mid$(varGlobal, 1, 6) & "' and a.numeromov = '" & Mid$(varGlobal, 7, 6) & "'"
    rsCriterio.Open SqlCriterio, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCriterio.RecordCount > 0 Then
        compoeControlesForm
    End If
    rsCriterio.Close
    Set rsCriterio = Nothing
End Sub

Private Sub compoeControlesForm()
    txtEmprestimo(0) = rsCriterio.Fields(0) 'Chapa do colaborador
    txtEmprestimo(1) = rsCriterio.Fields(1) 'Nome do colaborador
    DTPicker1.Value = rsCriterio.Fields(2) 'Data do emprestimo
    txtEmprestimo(8) = rsCriterio.Fields(4) 'Nome do colaborador emprestou o produto
End Sub

Private Sub BloqControl()
    Dim X As Integer
    For X = 0 To txtEmprestimo.Count - 1
        txtEmprestimo(X).Enabled = False
    Next
    For X = 0 To cmdEmp.Count - 1
        cmdEmp(X).Enabled = False
    Next
    For X = 0 To cmdEmprestimo.Count - 1
        cmdEmprestimo(X).Enabled = False
    Next
    DTPicker1.Enabled = False
    chkEmprestimo.Enabled = False
    ListView1.Enabled = False
    ListView2.Enabled = False
End Sub


Private Sub GravaTMov()
    Dim rsGravaTMov As New ADODB.Recordset
    Dim SqlGravaTMov As String
    Dim vValor As Double
   
    SqlGravaTMov = "Select A.CODCOLIGADA,A.IDMOV,A.CODFILIAL,A.CODLOC,A.CODCFO,A.CODCFONATUREZA,A.NUMEROMOV,A.SERIE,A.CODTMV,A.TIPO,A.STATUS,A.MOVIMPRESSO,A.DOCIMPRESSO,A.FATIMPRESSA,A.DATAEMISSAO,A.COMISSAOREPRES,A.VALORBRUTO,A.VALORLIQUIDO,A.VALOROUTROS,A.PERCCOMISSAO,A.PESOLIQUIDO," & _
    "A.PESOBRUTO,A.CODMOEVALORLIQUIDO,A.DATAMOVIMENTO,A.GEROUFATURA,A.CODCFOAUX,A.CODVEN1,A.CODVEN2,A.PERCCOMISSAOVEN2,A.CODCOLCFO,A.CODCOLCFONATUREZA,A.CODUSUARIO,A.GERADOPORLOTE,A.STATUSEXPORTCONT,A.GEROUCONTATRABALHO,A.GERADOPORCONTATRABALHO,A.HORULTIMAALTERACAO," & _
    "A.INDUSOOBJ,A.CONTABILIZADOPORTOTAL,A.INTEGRADOBONUM,A.FLAGPROCESSADO,A.ABATIMENTOICMS,A.USUARIOCRIACAO,A.DATACRIACAO,A.STSEMAIL,A.VALORBRUTOINTERNO,A.VINCULADOESTOQUEFL,A.VALORDESCCONDICIONAL,A.VALORDESPCONDICIONAL,A.CONTORCAMENTOANTIGO,A.SEQUENCIALESTOQUE," & _
    "A.INTEGRADOAUTOMACAO,A.INTEGRAAPLICACAO,A.DATALANCAMENTO,A.EXTENPORANEO,A.RECIBONFESTATUS,A.IDMOVCFO,A.VALORMERCADORIAS,A.USARATEIOVALORFIN,A.CODCOLCFOAUX,A.VRBASEINSSOUTRAEMPRESA,A.VALORBRUTOORIG,A.VALORLIQUIDOORIG,A.VALOROUTROSORIG,A.RECCREATEDBY,A.RECCREATEDON," & _
    "A.RECMODIFIEDBY,A.RECMODIFIEDON from tmov as a where a.idmov = '" & vIDMov & "'"
    rsGravaTMov.Open SqlGravaTMov, cnBancoSAP, adOpenKeyset, adLockOptimistic
 
    vValor = Format(Text1.Text, "#,##0.00;(#,##0.00)")
    
    If rsGravaTMov.RecordCount = 0 Then
        rsGravaTMov.AddNew
        rsGravaTMov.Fields(0) = vCodColigadaRM
        rsGravaTMov.Fields(1) = vIDMov
        rsGravaTMov.Fields(2) = 1
        rsGravaTMov.Fields(3) = vLocalEstoque
        rsGravaTMov.Fields(4) = "000001"
        rsGravaTMov.Fields(5) = "000001"
        rsGravaTMov.Fields(6) = Format(vNumeromov, "000000")
        rsGravaTMov.Fields(7) = vSerie
        rsGravaTMov.Fields(8) = "2.2.15"
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
    
    
 '   vValor = Format(Text1.Text, "#,##0.00;(#,##0.00)")
 '   SqlGravaTMov = "Insert into tmov(" & _
 '       "CODCOLIGADA,IDMOV,CODFILIAL,CODLOC,CODCFO,CODCFONATUREZA,NUMEROMOV,SERIE,CODTMV,TIPO,STATUS,MOVIMPRESSO,DOCIMPRESSO,FATIMPRESSA,DATAEMISSAO,COMISSAOREPRES,VALORBRUTO,VALORLIQUIDO,VALOROUTROS,PERCCOMISSAO,PESOLIQUIDO," & _
 '       "PESOBRUTO,CODMOEVALORLIQUIDO,DATAMOVIMENTO,GEROUFATURA,CODCFOAUX,CODVEN1,CODVEN2,PERCCOMISSAOVEN2,CODCOLCFO,CODCOLCFONATUREZA,CODUSUARIO,GERADOPORLOTE,STATUSEXPORTCONT,GEROUCONTATRABALHO,GERADOPORCONTATRABALHO,HORULTIMAALTERACAO," & _
 '       "INDUSOOBJ,CONTABILIZADOPORTOTAL,INTEGRADOBONUM,FLAGPROCESSADO,ABATIMENTOICMS,USUARIOCRIACAO,DATACRIACAO,STSEMAIL,VALORBRUTOINTERNO,VINCULADOESTOQUEFL,VALORDESCCONDICIONAL,VALORDESPCONDICIONAL,CONTORCAMENTOANTIGO,SEQUENCIALESTOQUE," & _
 '       "INTEGRADOAUTOMACAO,INTEGRAAPLICACAO,DATALANCAMENTO,EXTENPORANEO,RECIBONFESTATUS,IDMOVCFO,VALORMERCADORIAS,USARATEIOVALORFIN,CODCOLCFOAUX,VRBASEINSSOUTRAEMPRESA,VALORBRUTOORIG,VALORLIQUIDOORIG,VALOROUTROSORIG,RECCREATEDBY,RECCREATEDON," & _
 '       "RECMODIFIEDBY,RECMODIFIEDON) " & _
 '       "Values(" & CInt(vCodColigadaRM) & "," & vIDMov & ",1,'" & Mid$(txtEmprestimo(7).Text, 1, 4) & "','000001','000001','" & vNumeromov & "','" & vSerie & "','2.2.15','P','N',0,0," & _
 '       "0,'" & DTPicker1.Value & "',0," & vValor & "," & vValor & "," & vValor & ",0,0,0,'R$','" & DTPicker1.Value & "',0,'CXXXXXXXXXX','" & txtEmprestimo(0).Text & "','" & vCodVenRM & "',0,1,1,'" & vCodUsuarioRM & "',0,0,0,0,'" & Time & "',0,0,0,0,0,'" & vCodUsuarioRM & "','" & DTPicker1.Value & "',0," & vValor & ",0,0," & _
 '       "0,0,'" & vSequencialEstoque & "',0,'T','" & DTPicker1.Value & "',0,0,539,0,0,0,0,0,0,0,'" & vCodUsuarioRM & "','" & DTPicker1.Value & "','" & vCodUsuarioRM & "','" & DTPicker1.Value & "')"
            
''        Text2.Text = SqlGravaTMov
''        Exit Sub
'    rsGravaTMov.Open SqlGravaTMov, cnBancoSAP
    
    
    
End Sub

Private Sub GravaTitMMov(vSequencialItens As Integer)
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
        rsGravaTitMMov.Fields(4) = ListView2.SelectedItem.ListSubItems.Item(7)
        rsGravaTitMMov.Fields(5) = ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(6) = ListView2.SelectedItem.ListSubItems.Item(6) / ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(7) = 0
        rsGravaTitMMov.Fields(8) = DTPicker1.Value
        rsGravaTitMMov.Fields(9) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaTitMMov.Fields(10) = ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(11) = 1
        rsGravaTitMMov.Fields(12) = ListView2.SelectedItem.ListSubItems.Item(6) / ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(13) = ListView2.SelectedItem.ListSubItems.Item(6) / ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(14) = 0
        rsGravaTitMMov.Fields(15) = ListView2.SelectedItem.ListSubItems.Item(2)
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
        rsGravaTitMMov.Fields(35) = vLocalEstoque 'local de estoque
        rsGravaTitMMov.Fields(36) = 0
        rsGravaTitMMov.Fields(37) = ListView2.SelectedItem.ListSubItems.Item(6) / ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(38) = ListView2.SelectedItem.ListSubItems.Item(6) / ListView2.SelectedItem.ListSubItems.Item(2)
        rsGravaTitMMov.Fields(39) = 0
        rsGravaTitMMov.Fields(40) = ListView2.SelectedItem.ListSubItems.Item(2)
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
End Sub

Private Sub GravaTprdLoc()
    Dim rsGravaTprdLoc As New ADODB.Recordset
    Dim SqlGravaTprdLoc As String

    SqlGravaTprdLoc = "UPDATE TPRDLOC set SALDOFISICO2 = SALDOFISICO2-'" & ListView2.SelectedItem.ListSubItems.Item(2) & "' where codcoligada = '" & vCodColigadaRM & "' and CODLOC = " & vLocalEstoque & " AND CODFILIAL = 1 AND IDPRD = '" & ListView2.SelectedItem.ListSubItems.Item(7) & "'"
    rsGravaTprdLoc.Open SqlGravaTprdLoc, cnBancoSAP

    SqlGravaTprdLoc = "UPDATE TPRDLOC set SALDOFINANCEIRO2 = SALDOFISICO2*CUSTOMEDIO where codcoligada = '" & vCodColigadaRM & "' and CODLOC = " & vLocalEstoque & " AND CODFILIAL = 1 AND IDPRD = '" & ListView2.SelectedItem.ListSubItems.Item(7) & "'"
    rsGravaTprdLoc.Open SqlGravaTprdLoc, cnBancoSAP


End Sub
