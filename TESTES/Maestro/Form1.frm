VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   6120
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   1095
      Left            =   4320
      Picture         =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin MAESTRO.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   255
      MPTR            =   1
      MICON           =   "Form1.frx":0CCA
      PICN            =   "Form1.frx":0CE6
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Msgbox myskin
End Sub

Private Sub Form_Load()
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    OrganizaForm
    
    chameleonButton1.BackColor = Principal.Skin1.WindowColor
    chameleonButton1.BackOver = Principal.Skin1.WindowColor
    chameleonButton1.MaskColor = Principal.Skin1.WindowColor
    
'    chameleonButton1.BackColor = Principal.Ribbon.BackColor
'    chameleonButton1.BackOver = Principal.Ribbon.BackColor
'    chameleonButton1.MaskColor = Principal.Ribbon.BackColor
    'OrganizaControles
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    'OrganizaControles
End Sub

Private Function OrganizaForm()
    Me.Move 0, 0, Principal.ScaleWidth - 50, Principal.ScaleHeight - 50
End Function

Private Function OrganizaControles()
    On Error Resume Next
    Frame1.Move 1920, 0, Me.ScaleWidth - 2000
    Frame3.Move 0, 0, Me.ScaleWidth - Me.ScaleWidth + 1800, Me.ScaleHeight
    DBGrid2.Move 1920, 1080, Me.ScaleWidth - 2000, Me.ScaleHeight - 1080
End Function
