VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   3360
      Top             =   120
   End
   Begin Project1.VistaProgress VistaProgress1 
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   397
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    VistaProgress1.Value = 0
End Sub

Private Sub Timer1_Timer()
    If VistaProgress1.Value <> VistaProgress1.Max Then
        VistaProgress1.Value = VistaProgress1.Value + 1
    Else
        Timer1.Enabled = False
        VistaProgress1.Value = 0
    End If
End Sub

