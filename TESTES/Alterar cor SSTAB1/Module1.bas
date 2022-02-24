Attribute VB_Name = "Module1"
Option Explicit

Public Function SSTabProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    SSTabProc = Form1.NewSSTabProc(Hwnd, uMsg, wParam, lParam)
    On Error GoTo 0
End Function
