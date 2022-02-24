VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Large Vista ProgressBar"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   240
      Top             =   1080
   End
   Begin Project1.VistaProgressLarge VistaProgressLarge1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
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

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 100, 100, Form1.Width * 0.06, Form1.Height * 0.06, 0
    Ret = GetWindowLong(Me.hwnd, -20)
    Ret = Ret Or &H80000
    SetWindowLong Me.hwnd, -20, Ret
    SetLayeredWindowAttributes Me.hwnd, vbBlack, 255, &H2
End Sub

Private Sub Timer1_Timer()
    If VistaProgressLarge1.Value <> 100 Then
        VistaProgressLarge1.Value = VistaProgressLarge1.Value + 1
    Else
        Timer1.Enabled = False
    End If
End Sub
