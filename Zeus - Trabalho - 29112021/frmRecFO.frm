VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmRecFO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receber FO"
   ClientHeight    =   2280
   ClientLeft      =   5220
   ClientTop       =   4050
   ClientWidth     =   5070
   Icon            =   "frmRecFO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      Picture         =   "frmRecFO.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmRecFO.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Frame Frame2 
         Caption         =   "Serviço "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txtCadastro 
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
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRecFO.frx":265E
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.OptionButton optCadastro 
         Caption         =   "Executar Serviço"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optCadastro 
         Caption         =   "Arquivar"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRecFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If optCadastro(0).Value = True Then
            ArquivaFO
            'Unload Me
        Else
            If ProcuraFCE = False Then
                Msgbox "FCE já cadastrada", vbInformation, "Zeus"
                txtCadastro = Format(GeraFCE, "000000") & ""
                Exit Sub
            Else
                txtCadastro = Format(txtCadastro, "000000") & ""
            End If
            varGlobal2 = txtCadastro
            frmFCE.Show 1
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    MarcaPosicaoLV
    If VerificaChecados = False Then
        mobjMsg.Abrir "Nenhuma FO foi marcada", Ok, critico, "ZEUS"
        Unload Me
    End If
    vListViewPrincipal.ListItems(Posicao).Selected = True
End Sub

Private Function VerificaChecados()
    Dim X As Integer, Y As Integer
    VerificaChecados = False
    Y = vListViewPrincipal.ListItems.Count
    For X = 1 To Y
        vListViewPrincipal.ListItems(X).Selected = True
        vListViewPrincipal.ListItems(X).EnsureVisible
        If vListViewPrincipal.ListItems.Item(X).Checked = True Then
            VerificaChecados = True
        End If
    Next
End Function

Private Sub MarcaPosicaoLV()
    Dim X As Integer, Y As Integer
    Y = vListViewPrincipal.ListItems.Count
    For X = 1 To Y
        If vListViewPrincipal.ListItems.Item(X).Selected = True Then
            vListViewPrincipal.ListItems.Item(X).Selected = True
            Exit For
        End If
    Next
    Posicao = X
End Sub

Private Sub Form_Load()
    carregarIconBotao
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub optCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        txtCadastro = ""
        Frame2.Enabled = False
        Label1.Enabled = False
    Case 1
        Frame2.Enabled = True
        Label1.Enabled = True
        txtCadastro = Format(GeraFCE, "000000") & ""
    End Select
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdCadastro(0), 0, 49 'Inserir
    carregaImagemBotao cmdCadastro(1), 1, 34 'Sair
End Sub

Private Function GeraFCE()
On Error GoTo Err
    Dim rsGeraFCE As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tbFCE order by fce Desc"
    rsGeraFCE.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraFCE.RecordCount > 0 Then
        GeraFCE = rsGeraFCE.Fields(0) + 1
    Else
        QualForm = "novafce"
        GeraFCE = NovoCodigo
    End If
    rsGeraFCE.Close
    Set rsGeraFCE = Nothing
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

Private Function ArquivaFO()
On Error GoTo Err
    Dim rsGravaFO As New ADODB.Recordset
    Dim sql As String
    
    sql = "Select * from tbfo"
    rsGravaFO.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    Y = frmPesqFO.ListView1.ListItems.Count
    For X = 1 To Y
        frmPesqFO.ListView1.ListItems(X).Selected = True
        frmPesqFO.ListView1.ListItems(X).EnsureVisible
        If frmPesqFO.ListView1.ListItems.Item(X).Checked = True Then
            While Not rsGravaFO.EOF
                If Val(frmPesqFO.ListView1.ListItems.Item(X)) = rsGravaFO.Fields(0) Then
                    rsGravaFO.Fields(2) = 3
                End If
                rsGravaFO.MoveNext
            Wend
            rsGravaFO.MoveFirst
        End If
    Next
    Msgbox "Ficha de Orçamento Arquivada com Sucesso!"
    If Not rsGravaFO.EOF Then rsGravaFO.Update
    rsGravaFO.Close
    Set rsGravaFO = Nothing
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

Private Sub txtCadastro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        If ProcuraFCE = False Then
            Msgbox "FCE já cadastrada", vbInformation, "Zeus"
            txtCadastro = Format(GeraFCE, "000000") & ""
            Exit Sub
        Else
            txtCadastro = Format(txtCadastro, "000000") & ""
        End If
        frmFCE.Show 1
    End If
End Sub

Private Function ProcuraFCE()
On Error GoTo Err
    ProcuraFCE = False
    Dim rsProcuraFCE As New ADODB.Recordset
    Dim SqlProcura As String
    SqlProcura = "Select top 1 * from tbFCE INNER JOIN TBFO ON TBFCE.FCE = TBFO.FCE where tbFCE.fce = " & Val(txtCadastro)
    rsProcuraFCE.Open SqlProcura, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProcuraFCE.RecordCount <= 0 Then
        ProcuraFCE = True
    End If
    rsProcuraFCE.Close
    Set rsProcuraFCE = Nothing
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

