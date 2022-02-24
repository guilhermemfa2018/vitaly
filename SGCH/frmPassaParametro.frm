VERSION 5.00
Begin VB.Form frmPassaParametro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetro do filtro"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Informe o valor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "frmPassaParametro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "frmPassaParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ValidaDados
    Unload Me
End Sub

Private Sub Form_Load()
    Frame1.Caption = "Filtro: " & frmFiltro.Combo1.Text
End Sub

Private Sub ValidaDados()
    If IsDate(Text1.Text) Then
        Text1.Text = Format(Text1, "yyyy-mm-dd")
        vAlteraLike = Text1.Text
    Else
        vAlteraLike = Text1.Text & "%"
    End If
    
    If IsDate(Text2.Text) Then
        Text2.Text = Format(Text2, "yyyy-mm-dd")
        vAlteraLike2 = Text2.Text
    Else
        vAlteraLike2 = Text2.Text & "%"
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        If Text2.Visible = False Then
            ValidaDados
            Unload Me
        End If
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        ValidaDados
        Unload Me
    End If
End Sub


