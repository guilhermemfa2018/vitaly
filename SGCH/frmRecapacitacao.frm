VERSION 5.00
Begin VB.Form frmRecapacitacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recapacitação"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frmRecapacitacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmRecapacitacao.frx":0CCA
         Left            =   120
         List            =   "frmRecapacitacao.frx":0CE0
         TabIndex        =   1
         Text            =   "Motivo-01"
         Top             =   240
         Width           =   7815
      End
   End
   Begin SGCH.chameleonButton cmdNovoCol 
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   960
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
      MICON           =   "frmRecapacitacao.frx":0D26
      PICN            =   "frmRecapacitacao.frx":0D42
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
      Top             =   960
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
      MICON           =   "frmRecapacitacao.frx":1A1C
      PICN            =   "frmRecapacitacao.frx":1A38
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
Attribute VB_Name = "frmRecapacitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNovoCol_Click(Index As Integer)
    Select Case Index
    Case 0
        GravarDados
        Unload Me
    Case 2
        If MsgBox("Deseja sair da tela de cadastro de recapacitação?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Pesquisa = "0"
            Unload Me
            Set frmRecapacitacao = Nothing
        End If
    End Select
End Sub

Private Sub GravarDados()
    If VerificaInativo = False Then Exit Sub
    VerificaInativo
    Dim rsRecap As New ADODB.Recordset
    Dim SqlRecap As String
    
    Dim rsNovaProg As New ADODB.Recordset
    Dim SqlNovaProg As String
    
    Dim rsConcluir As New ADODB.Recordset
    Dim SqlConcluir As String
    
    
    Dim X As Integer, novoID As Integer, vCodTrei As Integer
    Dim vCPF As String
    
    cnBanco.BeginTrans
    
    If chamaForm.Caption = "Recapacitação" Then
        For X = 1 To MeuLV.ListView1.ListItems.Count
            MeuLV.ListView1.ListItems.Item(X).Selected = True
            If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
                vCPF = MeuLV.ListView1.ListItems.Item(X)
                vCodTrei = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
                SqlRecap = "select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and status = 'Concluido' and situacao in( 'Aprovado com restrição','Reprovado') and cpf = '" & vCPF & "' and codtreinamento = '" & vCodTrei & "' order by id"
                rsRecap.Open SqlRecap, cnBanco, adOpenKeyset, adLockOptimistic
                rsRecap.Fields(4) = "N" 'Ativo
                rsRecap.Fields(6) = "Recapacitação" 'Status
                rsRecap.Fields(10) = Combo1.Text 'Motivo recapacitação
        
                SqlNovaProg = "select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' order by id desc"
                rsNovaProg.Open SqlNovaProg, cnBanco, adOpenKeyset, adLockOptimistic
                novoID = rsNovaProg.Fields(5) + 1
                rsNovaProg.AddNew
                rsNovaProg.Fields(0) = rsRecap.Fields(0) '
                rsNovaProg.Fields(1) = rsRecap.Fields(1)
                rsNovaProg.Fields(2) = rsRecap.Fields(2)
                rsNovaProg.Fields(4) = "S"
                rsNovaProg.Fields(5) = novoID
                rsNovaProg.Fields(6) = "Pendente"
                rsNovaProg.Fields(7) = 0
                rsNovaProg.Fields(12) = rsRecap.Fields(12)
                rsNovaProg.Fields(14) = vCodcoligada 'Codigo da coligada
        
                rsNovaProg.Update
                rsNovaProg.Close
                Set rsNovaProg = Nothing
        
                rsRecap.Update
                rsRecap.Close
                Set rsRecap = Nothing
            End If
        Next
    Else
'APROVAR DIRETO
        For X = 1 To MeuLV.ListView1.ListItems.Count
            MeuLV.ListView1.ListItems.Item(X).Selected = True
            If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
                
                vCPF = MeuLV.ListView1.ListItems.Item(X)
                vCodTrei = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
                
                SqlRecap = "select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and status = 'Concluido' and situacao = 'Aprovado com restrição' and cpf = '" & vCPF & "' and codtreinamento = '" & Val(vCodTrei) & "' order by id"
                rsRecap.Open SqlRecap, cnBanco, adOpenKeyset, adLockOptimistic
                rsRecap.Fields(6) = "Aprovado" 'Status
                rsRecap.Fields(10) = Combo1.Text 'Motivo recapacitação
                
                SqlConcluir = "select * from tbColaboradoresCur where codcoligada = '" & vCodcoligada & "' and cpf ='" & vCPF & "' and codtreinamento ='" & vCodTrei & "'and origem = 'SR'"
                rsConcluir.Open SqlConcluir, cnBanco, adOpenKeyset, adLockOptimistic
                rsConcluir.Fields(3) = "SA" ' SA - Sistema Aprovado
                rsConcluir.Update
                rsConcluir.Close
                Set rsConcluir = Nothing
                
                rsRecap.Update
                rsRecap.Close
                Set rsRecap = Nothing
            End If
        Next
    
    End If
    MsgBox "Os dados da Recapacitação foram salvos com sucesso", vbInformation, "SGCH"
    cnBanco.CommitTrans
End Sub

Private Function VerificaInativo()
    VerificaInativo = False
    Dim V As Integer
    For V = 1 To MeuLV.ListView1.ListItems.Count
        MeuLV.ListView1.ListItems.Item(V).Selected = True
        If MeuLV.ListView1.ListItems.Item(V).Checked = True Then
            If MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) <> "-" Then
                MsgBox "Existem colaboradores selecionados que ja foram recapacitados", vbCritical, "SGCH"
                Exit Function
            End If
        End If
    Next
    VerificaInativo = True
End Function

Private Function VerificaStatus()
    VerificaStatus = False
    Dim V As Integer
    For V = 1 To MeuLV.ListView1.ListItems.Count
        MeuLV.ListView1.ListItems.Item(V).Selected = True
        If MeuLV.ListView1.ListItems.Item(V).Checked = True Then
            If MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = "Reprovado" Then
                Exit Function
            End If
        End If
    Next
    VerificaStatus = True
End Function

Private Sub Form_Activate()
    If Pesquisa = "novo" Then
        If VerificaStatus = False Then
            MsgBox "Colaboradores REPROVADOS devem ser recapacitados", vbCritical, "SGCH"
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    If Pesquisa = "novo" Then
        chamaForm.Caption = "Aprovar"
        chamaForm.Frame1.Caption = "Selecione um motivos para APROVAÇÃO "
        VerificaStatus
    End If
    If Pesquisa <> "novo" Then
        chamaForm.Caption = "Recapacitação"
        chamaForm.Frame1.Caption = "Selecione um motivos para RECAPACITAÇÃO "
    End If
End Sub
