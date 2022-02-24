VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   1125
   ClientTop       =   1575
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   12645
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   1095
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "remover"
      Height          =   1095
      Left            =   1080
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton CmdClearVb5 
      Caption         =   "VB5"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CmdClearVB4 
      Caption         =   "VB4"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Sub VB5ClearDebug()
Dim parent_hwnd As Long

    parent_hwnd = FindWindow(vbNullString, "Immediate")
    If parent_hwnd = 0 Then Exit Sub

    SetFocusAPI parent_hwnd
    SendKeys "^{HOME}+^{END}^{BREAK}{DEL}{F5}", False
End Sub

Public Sub VB4ClearDebug()
Dim parent_hwnd As Long

    parent_hwnd = FindWindow(vbNullString, "Debug Window")
    If parent_hwnd = 0 Then Exit Sub

    SetFocusAPI parent_hwnd
    SendKeys "^{HOME}+^{END}{F5}{DEL}{F5}", True
End Sub

Private Sub CmdClearVB4_Click()
    VB4ClearDebug
End Sub

Private Sub CmdClearVb5_Click()
    VB5ClearDebug
End Sub

Private Sub Command2_Click()
    Unload Command1(1)
End Sub

Private Sub Command3_Click()
    Load Command1(1)
    With Command1(1)
        .Visible = True
        .Top = 60
        .Left = 1000
        .Width = 2535
        .Height = 1095
        .Caption = "TESTE"
        .Tag = "TESTE"
        .BackColor = &HB7B7B7
        .ZOrder (0)
    End With
End Sub
