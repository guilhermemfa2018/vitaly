VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Vista Progressbar"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Turn to red"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
   End
   Begin Project1.VistaProg VistaProg1 
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   397
   End
   Begin VB.Label Label1 
      Caption         =   "The Progressbar will be red if the value is beyond or equals to 80%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
VistaProg1.Value = 85
End Sub
