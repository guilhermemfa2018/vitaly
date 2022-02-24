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
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1560
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
      MICON           =   "frmRecFO.frx":0CCA
      PICN            =   "frmRecFO.frx":0CE6
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
      TabIndex        =   5
      Top             =   1560
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
      MICON           =   "frmRecFO.frx":19C0
      PICN            =   "frmRecFO.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
            OleObjectBlob   =   "frmRecFO.frx":26B6
            TabIndex        =   7
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
                txtcadastro = Format(GeraFCE, "000000") & ""
                Exit Sub
            Else
                txtcadastro = Format(txtcadastro, "000000") & ""
            End If
            varGlobal2 = txtcadastro
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
    MeuLV.ListView1.ListItems(Posicao).Selected = True
End Sub


Private Function VerificaChecados()
    Dim X As Integer, Y As Integer
    VerificaChecados = False
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        MeuLV.ListView1.ListItems(X).Selected = True
        MeuLV.ListView1.ListItems(X).EnsureVisible
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            VerificaChecados = True
        End If
    Next
End Function

Private Sub MarcaPosicaoLV()
    Dim X As Integer, Y As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            MeuLV.ListView1.ListItems.Item(X).Selected = True
            Exit For
        End If
    Next
    Posicao = X
End Sub

Private Sub Form_Load()
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
        txtcadastro = ""
        Frame2.Enabled = False
        Label1.Enabled = False
        'txtCadastro.Enabled = False
    Case 1
        Frame2.Enabled = True
        Label1.Enabled = True
        txtcadastro = Format(GeraFCE, "000000") & ""
    End Select
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
            txtcadastro = Format(GeraFCE, "000000") & ""
            Exit Sub
        Else
            txtcadastro = Format(txtcadastro, "000000") & ""
        End If
        frmFCE.Show 1
    End If
End Sub

Private Function ProcuraFCE()
On Error GoTo Err
    ProcuraFCE = False
    Dim rsProcuraFCE As New ADODB.Recordset
    Dim SqlProcura As String
    SqlProcura = "Select top 1 * from tbFCE INNER JOIN TBFO ON TBFCE.FCE = TBFO.FCE where tbFCE.fce = " & Val(txtcadastro)
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

