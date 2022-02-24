VERSION 5.00
Begin VB.Form testForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Handling events for arrays of objects"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "testForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   6
      Left            =   5760
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   5
      Left            =   4920
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   3960
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   3000
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   2160
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   1200
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   360
      Tag             =   "teste de tag"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"testForm.frx":000C
      Height          =   1170
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   5385
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "testForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' We must implement the widget event interface, so that we can handle its events.
Implements IWidgetEvents

' Define 6 widgets. Note that we are unable to use WithEvents.
Private myWidget(6) As Widget
'

Private Sub Form_Load()
    Dim i As Long
    
    ' Create 6 widgets and tell them to send us events.
    For i = 0 To 6
        Set myWidget(i) = New Widget
        Set myWidget(i).Callback = Me
        
    Next i
End Sub
'

Private Sub Image_Click(Index As Integer)
    If Index = 0 Then
        
    End If
    myWidget(Index).FireDummyEvent Index
    
End Sub

' The dummy event handler.
Private Sub IWidgetEvents_DummyEvent(ByVal DummyParameter As Long)
    MsgBox "You clicked " & DummyParameter
    
    
End Sub
'


Private Sub IMyEventInterface_MyEvent(AnyParams)
    ' our event handler
End Sub

'Private Sub testButton_Click(Index As Integer)
'    ' Fire the dummy event
'    myWidget(Index).FireDummyEvent Index
'End Sub
'
