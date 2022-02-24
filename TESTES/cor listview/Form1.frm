VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Default         =   -1  'True
      Height          =   735
      Left            =   9480
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1455
      Left            =   10680
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" _
(ByVal vKey As Long) As Integer
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Sub Form_Load()
Timer1.Interval = 100
End Sub
Private Sub Timer1_Timer()
If GetAsyncKeyState(VK_LBUTTON) Then
Label2.Caption = "Left Click"
ElseIf GetAsyncKeyState(VK_RBUTTON) Then
Label2.Caption = "Right Click"
Else
Label2.Caption = ""
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = Button
End Sub

