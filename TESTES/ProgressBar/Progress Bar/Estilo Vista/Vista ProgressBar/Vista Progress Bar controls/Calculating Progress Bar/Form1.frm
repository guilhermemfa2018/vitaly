VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Without that bloody flicker "
   ClientHeight    =   3225
   ClientLeft      =   9705
   ClientTop       =   7545
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin Project1.VistaProgress VistaProgress1 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disable"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable"
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
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   975
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
VistaProgress1.Enable = True
End Sub

Private Sub Command2_Click()
VistaProgress1.Enable = False
End Sub

Private Sub Command3_Click()
VistaProgress1.Pause = True
End Sub

Private Sub Command4_Click()
VistaProgress1.Pause = False
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 100, 100, Form1.Width * 0.065, Form1.Height * 0.07, 0
    Ret = GetWindowLong(Me.hwnd, -20)
    Ret = Ret Or &H80000
    SetWindowLong Me.hwnd, -20, Ret
    SetLayeredWindowAttributes Me.hwnd, vbBlack, 255, &H2
End Sub

