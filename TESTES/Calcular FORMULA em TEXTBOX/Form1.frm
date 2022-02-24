VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   32
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   720
      TabIndex        =   31
      Top             =   4200
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      DragIcon        =   "Form1.frx":0000
      Height          =   375
      Left            =   315
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variáveis "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2055
         Left            =   7320
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "Form1.frx":101F
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdCalc1 
         Caption         =   "Calcular"
         Height          =   495
         Index           =   16
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   6855
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL9"
         Height          =   495
         Index           =   8
         Left            =   4080
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL8"
         Height          =   495
         Index           =   7
         Left            =   4080
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL7"
         Height          =   495
         Index           =   6
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL6"
         Height          =   495
         Index           =   5
         Left            =   2160
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL5"
         Height          =   495
         Index           =   4
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL4"
         Height          =   495
         Index           =   3
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL3"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL2"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "VARIAVEL1"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "-"
         Height          =   495
         Index           =   13
         Left            =   6000
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "*"
         Height          =   495
         Index           =   11
         Left            =   6000
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   ")"
         Height          =   495
         Index           =   14
         Left            =   6600
         TabIndex        =   4
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "/"
         Height          =   495
         Index           =   12
         Left            =   6600
         TabIndex        =   3
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "+"
         Height          =   495
         Index           =   10
         Left            =   6600
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "("
         Height          =   495
         Index           =   9
         Left            =   6000
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblResult 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   18
         Top             =   2640
         Width           =   7335
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Configuração de conexão DB RM Sistemas "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3000
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   23
         Text            =   "vigamax"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "sa"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   21
         Text            =   "zeus"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Text            =   "srv1002\corporerm"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      Caption         =   "MT_PINTURA(0,4)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   840
      TabIndex        =   33
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VARIAVEL1  As Double
Dim VARIAVEL2 As Double
Dim VARIAVEL3 As Double
Dim VARIAVEL4 As Double
Dim VARIAVEL5 As Double
Dim VARIAVEL6 As Double
Dim VARIAVEL7 As Double
Dim VARIAVEL8 As Double
Dim VARIAVEL9 As Double
Public rsLocal As New ADODB.Recordset
Public Sqlp As String
Public Sqlp1 As String

Private Sub cmdCalc_Click(Index As Integer)
    Text1 = Text1 + cmdCalc(Index).Caption + " "
    Text1.SelStart = Len(Text1.Text)
    Text1.SetFocus
    
    Text8 = Text8 + cmdCalc(Index).Caption + " "
    Text8.SelStart = Len(Text8.Text)
    
    Select Case Index
    Case 15 'CALCULA DIRETO
        lblResult = Text1.Text
    End Select
End Sub

Private Sub cmdCalc1_Click(Index As Integer)
    calcular
End Sub

Private Sub Command1_Click()
    If Text8.Visible = False Then
        Text8.Visible = True
        Text8.SetFocus
        Label1.Top = 4560
    Else
        Text8.Visible = False
        Text1.Text = Text8.Text
        calcular
        Text3.SetFocus
        Label1.Top = 4080
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text3.Text = ""
    Text8.Text = ""
End Sub

Private Sub Form_Load()
    VARIAVEL1 = 10
    VARIAVEL2 = 3
    VARIAVEL3 = 6
    VARIAVEL4 = 5
    VARIAVEL5 = 9
    VARIAVEL6 = 20
    VARIAVEL7 = 15
    VARIAVEL8 = 25
    VARIAVEL9 = 65
    Conectar
End Sub

Public Sub calcular()
On Error Resume Next
    Sqlp = "SELECT CONVERT(DECIMAL(18,3), " & Text1.Text & ") "
    executaCalculo (Sqlp)
    
    'Sqlp1 = "SELECT CONVERT(DECIMAL(18,3), 20 / 3 )"
    
    rsLocal.Open Sqlp1, cnBanco, adOpenKeyset, adLockReadOnly
    lblResult.Caption = rsLocal.Fields(0)
    Text3.Text = rsLocal.Fields(0)
    Debug.Print rsLocal.Fields(0)
    rsLocal.Close
End Sub

Public Function executaCalculo(vSQL As String)
    vSQL = Replace(vSQL, "VARIAVEL1", VARIAVEL1)
    vSQL = Replace(vSQL, "VARIAVEL2", VARIAVEL2)
    vSQL = Replace(vSQL, "VARIAVEL3", VARIAVEL3)
    vSQL = Replace(vSQL, "VARIAVEL4", VARIAVEL4)
    vSQL = Replace(vSQL, "VARIAVEL5", VARIAVEL5)
    vSQL = Replace(vSQL, "VARIAVEL6", VARIAVEL6)
    vSQL = Replace(vSQL, "VARIAVEL7", VARIAVEL7)
    vSQL = Replace(vSQL, "VARIAVEL8", VARIAVEL8)
    vSQL = Replace(vSQL, "VARIAVEL9", VARIAVEL9)
    Sqlp1 = vSQL
End Function
