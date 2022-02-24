VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemitirColaborador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demitir colaborador"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmDemitirColaborador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   1095
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   1935
      Begin VB.TextBox txtDemCol 
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
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Tag             =   "Registro do novo colaborador"
         ToolTipText     =   "Registro do novo colaborador"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Registro nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da recisão "
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   5535
      Begin VB.TextBox txtDemColaborador 
         Height          =   315
         Index           =   4
         Left            =   3240
         TabIndex        =   2
         Tag             =   "Orgão da homologação"
         ToolTipText     =   "Orgão da homologação"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtDemColaborador 
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   1
         Tag             =   "Nº da homologação"
         ToolTipText     =   "Nº da homologação"
         Top             =   600
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   115474433
         CurrentDate     =   40709
      End
      Begin VB.Label Label7 
         Caption         =   "Orgão da homologação:"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Nº da homologação:"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Data da demissão:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         TabIndex        =   17
         Top             =   1200
         Width           =   3375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Foto "
      Height          =   2415
      Index           =   0
      Left            =   5760
      TabIndex        =   14
      Top             =   120
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   1635
         TabIndex        =   15
         Top             =   240
         Width           =   1695
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   2175
            Left            =   0
            Top             =   -120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   3836
            Image           =   "frmDemitirColaborador.frx":0CCA
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
   Begin VB.Frame Frame1 
      Caption         =   "Informações do candidato "
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtDemColaborador 
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
         TabIndex        =   10
         Tag             =   "Matriz e cargo do colaborador"
         Text            =   "matriz - cargo"
         ToolTipText     =   "Matriz e cargo do colaborador"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtDemColaborador 
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
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtDemColaborador 
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
         Left            =   120
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Frame Frame10 
         Caption         =   "Média geral"
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
         Left            =   3720
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Demitido"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CPF nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label42 
         Caption         =   "Matriz/Cargo:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   2055
      End
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
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
      MICON           =   "frmDemitirColaborador.frx":0CE2
      PICN            =   "frmDemitirColaborador.frx":0CFE
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
      TabIndex        =   3
      Tag             =   "Confirmar"
      ToolTipText     =   "Confirmar"
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
      MICON           =   "frmDemitirColaborador.frx":19D8
      PICN            =   "frmDemitirColaborador.frx":19F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   6135
   End
End
Attribute VB_Name = "frmDemitirColaborador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsDemColaboradores As New ADODB.Recordset
Private SqlDemColaboradores As String
Private rsGravaDem As New ADODB.Recordset
Private sqlGravaDem As String
'Private rsLocal As New ADODB.Recordset

Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        GravarDados
    Case 2
        Unload Me
        Set frmDemitirColaborador = Nothing
    End Select
End Sub

Private Sub Form_Activate()
    SqlDemColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.codcolaborador,a.homologacaonum,a.datarecisao,a.homologacaoorgao,a.ativo from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'colaborador' and a.ativo = 'S' and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' order by a.cpf"
    rsDemColaboradores.Open SqlDemColaboradores, cnBanco, adOpenKeyset, adLockReadOnly
    If Not IsNull(rsDemColaboradores.Fields(6)) Then
        txtDemColaborador(2).Text = rsDemColaboradores.Fields(6)
        txtDemColaborador(4).Text = rsDemColaboradores.Fields(8)
        DTPicker1 = rsDemColaboradores.Fields(7)
        MsgBox "Colaborador já DEMITIDO", vbCritical, "Atenção"
        rsDemColaboradores.Close
        Set rsDemColaboradores = Nothing
        Unload Me
    ElseIf IsNull(rsDemColaboradores.Fields(9)) Or rsDemColaboradores.Fields(9) = "N" Then
        MsgBox "Colaboradores não ativos não podem ser DEMITIDOS", vbCritical, "Atenção"
        rsDemColaboradores.Close
        Set rsDemColaboradores = Nothing
        Unload Me
    Else
        rsDemColaboradores.Close
        Set rsDemColaboradores = Nothing
    End If
End Sub

Private Sub Form_Load()
    ResultPesq
End Sub

Private Sub ResultPesq()
    SqlDemColaboradores = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.foto,a.codcolaborador,a.homologacaonum from tbColaboradores as a inner join tbColaboradoreshist as b on a.codcoligada ='" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo Where a.tipo = 'colaborador' and a.ativo = 'S' and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' order by a.cpf"
    rsDemColaboradores.Open SqlDemColaboradores, cnBanco, adOpenKeyset, adLockReadOnly
    If rsDemColaboradores.RecordCount > 0 Then
        CompoeControles
    Else
        MsgBox MeuLV.cmdconsulta(9).ToolTipText & " não encontrado"
    End If
    rsDemColaboradores.Close
    Set rsDemColaboradores = Nothing
End Sub

Private Sub CompoeControles()
On Error GoTo TrataErro1
    txtDemColaborador(0).Text = Mid$(varGlobal, 1, 11)
    txtDemColaborador(1).Text = rsDemColaboradores.Fields(1) 'MeuLV.ListView1.SelectedItem.ListSubItems.Item(1)
    Label41 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(5)
    txtDemCol(1).Enabled = False
    txtDemCol(1) = rsDemColaboradores(5)
    
    If Val(Label41) < MediaGlobal And Val(Label41) >= vAprovadoRest Then
        Label41.ForeColor = &H40C0&
    ElseIf Val(Label41) < vAprovadoRest Then
        Label41.ForeColor = &HC0&
    ElseIf Val(Label41) >= MediaGlobal Then
        Label41.ForeColor = &H8000&
        Frame4.Enabled = False
        Combo1.Enabled = False
    End If
    txtDemColaborador(3) = Format(rsDemColaboradores.Fields(2), "000000") & "-" & rsDemColaboradores.Fields(3)
    Label53.Caption = rsDemColaboradores.Fields(4)
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Sub GravarDados()
    If ValidaCampo = False Then Exit Sub
    If MsgBox("Confirma exclusão do " & LegendaExc & " selecionado?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
        sqlGravaDem = "UPDATE tbColaboradores set ativo = 'N',datarecisao = '" & Format(DTPicker1.Value, vFormatoDatetime) & "',homologacaonum = '" & txtDemColaborador(2) & "',homologacaoorgao = '" & txtDemColaborador(4) & "' Where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtDemColaborador(0) & "' and codcolaborador = '" & txtDemCol(1) & "'"
'        sqlGravaDem = "UPDATE tbColaboradores set ativo = 'N',datarecisao = '" & DTPicker1.Value & "',homologacaonum = '" & txtDemColaborador(2) & "',homologacaoorgao = '" & txtDemColaborador(4) & "' Where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtDemColaborador(0) & "' and codcolaborador = '" & txtDemCol(1) & "'"
        rsGravaDem.Open sqlGravaDem, cnBanco
        MsgBox "Colaborador demitido com sucesso", vbInformation, "SGCH"
        gravaLog "CPF: " & txtDemColaborador(0) & ", Registro: " & txtDemCol(1), "Nome: " & txtDemColaborador(1), "Média Geral: " & Label41 & ", Status: " & Label9
        excluiProgramacao
        excluiADP
        Unload Me
    End If
End Sub

Private Sub excluiProgramacao()
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    'Rotina de deletar se INTD estiver com status de Pendente
    sqlDeletar = "Delete from tbPendentesCur where tbPendentesCur.codcoligada ='" & vCodcoligada & "' and tbPendentesCur.cpf = '" & txtDemColaborador(0) & "' and status = 'Pendente' or tbPendentesCur.cpf = '" & txtDemColaborador(0) & "' and status = 'Agendado'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    sqlDeletar = "UPDATE tbINTD set ativo = 'N', observacao = '" & "A INTD foi CANCELADA pelo usuário: " & NomUsu & ", devido ao seguinte motivo apresentado: colaborador DEMITIDO" & "' , status = 'Cancelada' where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and codcolaborador = " & txtDemCol(1)
    rsDeletar.Open sqlDeletar, cnBanco
End Sub

Private Sub excluiADP()
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    'Rotina de deletar se Avaliação de desempenho em que o status da avaliação estiver
    'diferente de 'Concluido'
    sqlDeletar = "Delete from tbListaADP where tbListaADP.statusavaliacao <> 'Concluido' and tbListaADP.codcolaborador= '" & txtDemCol(1) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtDemColaborador(2).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtDemColaborador(2).Tag, vbInformation, "Atenção"
        Me.txtDemColaborador(2).SetFocus
        Exit Function
    End If
    If txtDemColaborador(4).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtDemColaborador(4).Tag, vbInformation, "Atenção"
        Me.txtDemColaborador(4).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function
