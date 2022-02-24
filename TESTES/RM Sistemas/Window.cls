VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public hwnd As Long
' GetWindow() Constants
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5

Public Function GetText() As String
    Dim sText As String * 255
    Dim tmpLen As Long
    tmpLen = GetWindowText(hwnd, sText, 500)
    GetText = Left(sText, tmpLen)
End Function

Public Function DesktopWindow() As Long
    DesktopWindow = GetDesktopWindow
End Function

Public Function SetText(sNewText As String) As Long
    SetText = SetWindowText(hwnd, sNewText)
End Function

Public Function GetChildWindow(hwndParent As Long)
    GetChildWindow = GetWindow(hwndParent, GW_CHILD)
End Function

Public Function NextWindow(hwnd As Long)
    NextWindow = GetNextWindow(hwnd, GW_HWNDNEXT)
End Function

Public Function GetWindowClassName() As String
    Dim sClass As String * 255
    Dim tmpLen As Long
    tmpLen = GetClassName(hwnd, sClass, 255)
    GetWindowClassName = Left(sClass, tmpLen)
End Function

Public Function Flash() As Long
    Flash = FlashWindow(hwnd, 1)
End Function

