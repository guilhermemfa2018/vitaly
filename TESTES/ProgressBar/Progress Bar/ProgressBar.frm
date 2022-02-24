VERSION 5.00
Object = "{6984D37E-D788-11D2-94CB-0080AD717E3A}#1.0#0"; "TrueProgressB.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin TrueProgressB.ProgressB ProgressB1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      ColorFrom       =   12648384
      ColorTo         =   16384
      Segmented       =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   5760
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
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
    ProgressB1.Value = 0
End Sub

Private Sub Timer1_Timer()
    If ProgressB1.Value <> ProgressB1.Max Then
        ProgressB1.Value = ProgressB1.Value + 1
    Else
        Timer1.Enabled = False
        ProgressB1.Value = 0
    End If
End Sub
