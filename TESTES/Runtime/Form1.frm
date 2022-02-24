VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   4935
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   3960
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Script.AddCode Text1.Text
Script.Run "Command1_Click"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Script.Run "Command1_Move"
End Sub

Private Sub Form_Load()
Script.AddObject "Command1", Command1
Script.AddObject "Text2", Text2
Script.AddObject "Form1", Form1
Script.AddCode Text1.Text
End Sub

