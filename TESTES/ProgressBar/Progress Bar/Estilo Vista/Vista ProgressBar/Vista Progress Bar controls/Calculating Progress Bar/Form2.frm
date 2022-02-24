VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Windows Vista - Hacked by Bill Gates"
   ClientHeight    =   3000
   ClientLeft      =   7740
   ClientTop       =   7965
   ClientWidth     =   7485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7485
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   7455
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1695
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   7575
         Begin VB.CommandButton Command6 
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
            Left            =   6120
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
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
            Left            =   4800
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
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
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
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
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   5
         Top             =   280
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   -120
      TabIndex        =   1
      Top             =   2400
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
         Default         =   -1  'True
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
         Left            =   4960
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Close"
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
         Left            =   6300
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin Project1.VistaProgress VistaProg 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
   End
End
Attribute VB_Name = "Form2"
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
End
End Sub

Private Sub Command2_Click()
If Frame3.Visible = False Then
    Command2.Caption = "Back"
    Frame3.Visible = True
Else
    Command2.Caption = "Next"
    Frame3.Visible = False
End If
End Sub

Private Sub Command3_Click()
VistaProg.Enable = True
End Sub

Private Sub Command4_Click()
VistaProg.Enable = False
End Sub

Private Sub Command5_Click()
VistaProg.Pause = True
End Sub

Private Sub Command6_Click()
VistaProg.Pause = False
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 504, 500, Form1.Width * 0.094, Form1.Height * 0.06, 0
    Ret = GetWindowLong(Me.hwnd, -20)
    Ret = Ret Or &H80000
    SetWindowLong Me.hwnd, -20, Ret
    SetLayeredWindowAttributes Me.hwnd, vbBlack, 255, &H2
    
VistaProg.Enable = True


Label1.Caption = "The above progress bar is consist of these features: " & vbCrLf & vbCrLf & "Enable - enables animation" & vbCrLf _
                 & "Disable - disables the animation" & vbCrLf & "Pause - animation will be paused wherever it is " & vbCrLf _
                 & "Continue -  Continues the animation if its paused"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
