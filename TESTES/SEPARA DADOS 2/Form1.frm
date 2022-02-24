VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11940
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19125
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   19125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   12120
      TabIndex        =   15
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpa"
      Height          =   735
      Left            =   7800
      TabIndex        =   13
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   "ORDER "
      Height          =   1335
      Left            =   8400
      TabIndex        =   7
      Top             =   9480
      Width           =   9855
      Begin VB.TextBox Text 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "GROUP "
      Height          =   1335
      Left            =   8400
      TabIndex        =   6
      Top             =   8040
      Width           =   9855
      Begin VB.TextBox Text 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "WHERE "
      Height          =   1095
      Left            =   8400
      TabIndex        =   5
      Top             =   6840
      Width           =   9855
      Begin VB.TextBox Text 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FROM "
      Height          =   1455
      Left            =   8400
      TabIndex        =   4
      Top             =   5280
      Width           =   9855
      Begin VB.TextBox Text 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT "
      Height          =   3135
      Left            =   8400
      TabIndex        =   3
      Top             =   2040
      Width           =   9855
      Begin VB.TextBox Text 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   6855
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   1920
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Separa"
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Qtd. FROM:"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    QuantidadeAssociada Text1.Text
End Sub


Private Function QuantidadeAssociada(ByVal stringParaConsulta As String)
On Error GoTo Err
    Dim vPoints(7, 1) As String
    Dim K As Integer, Y As Integer, vInicio As Integer, vFim As Integer
    Dim ondeProdurar As String
    Dim oqueProcurar As String
    Dim vPosicao As Integer
    
    vPoints(0, 0) = "SELECT"
    vPoints(0, 1) = 1
    vPoints(1, 0) = "FROM"
    vPoints(1, 1) = 3
    vPoints(2, 0) = "WHERE"
    vPoints(2, 1) = 1
    vPoints(3, 0) = "GROUP"
    vPoints(3, 1) = 1
    vPoints(4, 0) = "ORDER"
    vPoints(4, 1) = 1
    vPoints(5, 0) = ";GO"
    vPoints(5, 1) = "1"
    vPoints(6, 0) = ";GO"
    vPoints(6, 1) = "1"
    vPoints(7, 0) = ";GO"
    vPoints(7, 1) = "1"
    vPosicao = 1
    For K = 0 To 5
        ondebuscar = stringParaConsulta
        oquebuscar = vPoints(K, 0)
        If vPoints(K, 1) > 1 Then
            For Y = 1 To vPoints(K, 1)
                vPosicao = InStr(vPosicao, ondebuscar, oquebuscar) + 1
            Next
            vPoints(K, 1) = vPosicao - 1
        Else
            vPoints(K, 1) = InStr(vPosicao, ondebuscar, oquebuscar)
            If vPoints(K, 1) <> 0 Then vPosicao = vPoints(K, 1)
        End If
        If K = 5 Then vPoints(K, 1) = 10000
    Next
    Y = 0
    For K = 0 To 5
        vInicio = vPoints(K, 1)
        If vPoints(K + 1, 1) > 0 Then vFim = vPoints(K + 1, 1) - 1
        If vFim = -1 Then vFim = 10000
        Text(Y).Text = Mid$(ondebuscar, vInicio, vFim - vInicio)
        Y = Y + 1
    Next
    Exit Function
Err:
    If K = 5 Then Exit Function
    If Err.Number = 9 Then Exit Function
    If Err.Number = 5 Then K = K + 1
    If vInicio = 0 And vFim = 0 Then Exit Function
    If vInicio = 0 Then
        vInicio = vPoints(K + 1, 1) - 1
        vFim = vPoints(K + 2, 1) - 1
        K = K + 1
    Else
        vFim = vPoints(K + 1, 1)
    End If
    Resume
End Function

Private Sub Command2_Click()
On Error Resume Next
    For K = 0 To 4
        Text(K).Text = ""
    Next
End Sub

