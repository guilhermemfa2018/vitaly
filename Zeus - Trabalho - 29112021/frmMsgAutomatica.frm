VERSION 5.00
Begin VB.Form frmMsgAutomatica 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   495
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2160
      Top             =   0
   End
   Begin ZEUS.VistaProgress VistaProgress1 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   397
   End
End
Attribute VB_Name = "frmMsgAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If vTimer = True Then
        Timer1.Enabled = True
        'Unload Me
        Me.Visible = True
    Else
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    If VistaProgress1.Value <> VistaProgress1.Max Then
        VistaProgress1.Value = VistaProgress1.Value + 1
    Else
        Timer1.Enabled = False
        VistaProgress1.Value = 0
        Unload Me
    End If
    
    If vTimer = False Then
        Unload Me
    End If
End Sub

Private Sub Timer2_Timer()
    Unload Me
End Sub
