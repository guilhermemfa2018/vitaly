VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Normal Vista Progress bar"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6000
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enable Timer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin Project1.VistaProgress ProgCrystal 
      Height          =   225
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   397
      Value           =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Text            =   "50"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Text            =   "100"
      Top             =   705
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max                Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type


Dim a As Single

Private Sub Command1_Click()
ProgCrystal.Value = Text2.Text
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Enable Timer" Then
    ProgCrystal.Value = 0
    Timer2.Enabled = True
    Command2.Caption = "Disable Timer"
Else
    Timer2.Enabled = False
    Command2.Caption = "Enable Timer"
End If
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 100, 100, Form1.Width * 0.06, Form1.Height * 0.06, 0
Ret = GetWindowLong(Me.hwnd, -20)
Ret = Ret Or &H80000
SetWindowLong Me.hwnd, -20, Ret
SetLayeredWindowAttributes Me.hwnd, vbBlack, 255, &H2
End Sub

Private Sub Text1_Change()
VistaProgress1.Max = Text1.Text
End Sub

Private Sub Timer2_Timer()
If ProgCrystal.Value <> ProgCrystal.Max Then
    ProgCrystal.Value = ProgCrystal.Value + 1
Else
    Timer2.Enabled = False
    Command2.Caption = "Enable Timer"
End If
End Sub
