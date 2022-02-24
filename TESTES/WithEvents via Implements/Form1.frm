VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6690
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   9420
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements Class1

Private m_Class1(1 To 3) As New Class1

Private Sub Class1_SomeEvent(ByVal Index As Long, ByVal Param As Long, RetVal As String)
    Print Index, Param, ;
    RetVal = Switch(Index = 1&, "vbLeftButton", _
                    Index = 2&, "vbMiddleButton", _
                    Index = 3&, "vbRightButton")
End Sub

Private Sub Form_Load()
    Dim i As Long

    For i = 1& To 3&
        Set m_Class1(i).Callback(i) = Me
    Next

    AutoRedraw = True
    Caption = "Simulating WithEvents for Arrays of Objects via Implements Demo"
    Print "Index", "Button", "Const"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Class1(Choose(Button, 1&, 3&, , 2&)).TriggerAnEvent Param:=Button
End Sub
