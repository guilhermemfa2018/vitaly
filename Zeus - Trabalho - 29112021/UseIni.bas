Attribute VB_Name = "UseIni"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetProfile(lpAppName As String, lpKeyName As String, lpDefault, lpFileName As String)
    Dim lpReturnString As String
    Dim nSize As Integer
    Dim Valid As Integer
    
    lpReturnString = Space(128)
    nSize = Len(lpReturnString)
    Valid = GetPrivateProfileString(ByVal lpAppName, ByVal lpKeyName, ByVal lpDefault, ByVal lpReturnString, ByVal nSize, ByVal lpFileName)
    GetProfile = Left(lpReturnString, Valid)
End Function

Public Sub WriteProfile(lpAppName As String, lpKeyName As String, lpString As String, lpFileName As String)
    Dim Valid As Integer
    
    Valid = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

Public Function GetProfileSection(lpAppName As String, lpFileName As String) As String
    Dim strReturnString As String
    Dim lSize As Long, lValid As Long
    
    strReturnString = Space(256)
    lSize = Len(strReturnString)
    lValid = GetPrivateProfileSection(ByVal lpAppName, ByVal strReturnString, ByVal lSize, ByVal lpFileName)
    GetProfileSection = Left(strReturnString, lValid)
End Function

Public Function GetValue(IniFileName As String, Section As String, Key As String, Optional ByVal defaultValue As String) As String
  On Error GoTo Hell
  Dim Value As String, retval As String, X As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(Section, Key, defaultValue, retval, Len(retval), IniFileName)
  GetValue = Trim(Left(retval, X))
Exit Function
Hell:
  GetValue = defaultValue
End Function




