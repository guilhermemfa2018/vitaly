VERSION 5.00
Begin VB.Form frmExcluiINTD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusão de INTD"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmExcluiINTD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Observação"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmExcluiINTD.frx":0CCA
         Left            =   120
         List            =   "frmExcluiINTD.frx":0CD1
         TabIndex        =   10
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações da INTD"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtDemINTD 
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
         TabIndex        =   3
         Text            =   "Registro"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtDemINTD 
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
         TabIndex        =   2
         Text            =   "INTD"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtDemINTD 
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
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Tag             =   "Matriz e cargo do colaborador"
         Text            =   "Nome"
         ToolTipText     =   "Matriz e cargo do colaborador"
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label42 
         Caption         =   "Colaborador:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "INTD nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin SGCH.chameleonButton cmdExcINTD 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   3480
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
      MICON           =   "frmExcluiINTD.frx":0CF0
      PICN            =   "frmExcluiINTD.frx":0D0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdExcINTD 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Tag             =   "Confirmar exclusão"
      ToolTipText     =   "Confirmar exclusão"
      Top             =   3480
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
      MICON           =   "frmExcluiINTD.frx":19E6
      PICN            =   "frmExcluiINTD.frx":1A02
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
Attribute VB_Name = "frmExcluiINTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsExcINTD As New ADODB.Recordset
Private SqlExcINTD As String
Private rsGravaExcINTD As New ADODB.Recordset
Private sqlGravaExcINTD As String

Private Sub cmdExcINTD_Click(Index As Integer)
    Select Case Index
    Case 0
        excluirDadosINTD
        'gravaLog "CPF: " & txtNovoColaborador(0) & ", Registro: " & txtNovoCol(1), "Nome: " & txtNovoColaborador(1), "Média Geral: " & Label41 & ", Status: " & Label9
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = "N" Then
        MsgBox "Esta INTD já está CANCELADA", vbCritical, "SGCH"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CompoeControles
End Sub

Private Sub CompoeControles()
On Error GoTo TrataErro1
    txtDemINTD(0).Text = varGlobal
    txtDemINTD(1).Text = MeuLV.ListView1.SelectedItem.ListSubItems.Item(3)
    txtDemINTD(2) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Sub excluirDadosINTD()
    If ValidaCampo = False Then Exit Sub
    SqlExcINTD = "UPDATE tbINTD set ativo = 'N', observacao = '" & "A INTD foi CANCELADA pelo usuário: " & NomUsu & ", devido aa seguinte motivo apresentado: " & Combo1.Text & "' ,status = 'Cancelada' where codcoligada = '" & vCodColigada & "' and codINTD = " & Val(txtDemINTD(0))
    rsExcINTD.Open SqlExcINTD, cnBanco

    SqlExcINTD = "Delete from tbPendentesCur where codcoligada = '" & vCodColigada & "' and codINTD= '" & Val(txtDemINTD(0)) & "'"
    rsExcINTD.Open SqlExcINTD, cnBanco

    MsgBox "INTD CANCELADA com sucesso!", vbInformation, "SGCH"
    AtualizaListview
    Unload Me
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If Combo1 = "" Then
        MsgBox "Favor informar o campo " & Me.Combo1.Tag, vbCritical, "Atenção"
        Me.Combo1.SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = "Cancelada"
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK"
    Exit Sub
Err:
    MsgBox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

