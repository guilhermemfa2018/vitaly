Attribute VB_Name = "Help"
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal HWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

'definição de constantes usadas pela API

Public Const HELP_CONTEXT = &H1
Public Const HELP_QUIT = &H2
Public Const HELP_INDEX = &H3
Public Const HELP_HELPONHELP = &H4
Public Const HELP_SETINDEX = &H5
Public Const HELP_KEY = &H101
Public Const HELP_MULTIKEY = &H201

Public vHelp As String
