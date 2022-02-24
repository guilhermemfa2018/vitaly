Attribute VB_Name = "Module1"
Option Explicit
'PARA FECHAR O PROGRAMA PELO GERENCIADOR DE TAREFAS
Public Const MAX_PATH As Integer = 260
Type PROCESSENTRY32
     dwSize As Long
     cntUsage As Long
     th32ProcessID As Long
     th32DefaultHeapID As Long
     th32ModuleID As Long
     cntThreads As Long
     th32ParentProcessID As Long
     pcPriClassBase As Long
     dwFlags As Long
     szExeFile As String * MAX_PATH
End Type

Public Const TH32CS_SNAPPROCESS = &H2
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Global hSnapShot As Long
Global uProcess As PROCESSENTRY32
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Global rProcess As Long
Global tPID As Long
Global tMID As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long
Public Const PROCESS_TERMINATE = &H1
Global cescolhido As String
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Global hProcess As Long
Global lExitCode As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Type LUID: UsedPart As Long: IgnoredForNowHigh32BitPart As Long: End Type
Private Type OSVERSIONINFO: dwOSVersionInfoSize As Long: dwMajorVersion As Long: dwMinorVersion As Long: dwBuildNumber As Long: dwPlatformId As Long: szCSDVersion As String * 128: End Type
Private Type TOKEN_PRIVILEGES: PrivilegeCount As Long: TheLuid As LUID: Attributes As Long: End Type
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

'PARA MANIPULAÇÃO DO REGEDIT
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'ABAIXO DEIXA PROGRAMA NO SYSTRAY
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'NAO DEIXA O APLICATIVO ABRIR MAIS DE UMA VEZ
Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'ABAIXO MONITORA TEMPO DE OCIOSIDADE
Public Declare Function timeGetTime Lib "WINMM.DLL" () As Long
Private Type LASTINPUTINFO
  cbSize As Long
  dwTime As Long
End Type

Public Declare Function GetLastInputInfo Lib "user32.dll" (plii As LASTINPUTINFO) As Long
Public cnBanco As ADODB.Connection
Public controleAcesso As Integer
Public senhaUsu As String

Public Function fnIdleTime() As Long
    Dim lii As LASTINPUTINFO
    lii.cbSize = Len(lii)
    If (GetLastInputInfo(lii) > 0) Then
        fnIdleTime = (timeGetTime - lii.dwTime) \ 1000
    End If
End Function

'ABAIXO CONEXÃO COM O BANCO DE DADOS
Public Function Conectar()
On Error GoTo Err1
    If Form1.Text1.Text = "" Then GoTo Err1 'nome do servidor
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Password=" & Form1.Text4.Text & ";Persist Security Info=True;User ID=" & Form1.Text3.Text & ";Initial Catalog=" & Form1.Text2.Text & ";Data Source=" & Form1.Text1.Text
    Form1.Label20.Visible = False
    'Form1.Visible = False
    Exit Function
Err1:
    Form1.Label20.Visible = True
    Form1.WindowState = 0 ' normal
    Exit Function
End Function

'PARA FECHAR O PROGRAMA PELO GERENCIADOR DE TAREFAS
Function GetProcessIDByEXEName(ByVal EXEName As String) As Long
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim R As Long, lStrtemp As String
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = -1 Then Exit Function
    uProcess.dwSize = Len(uProcess)
    R = ProcessFirst(hSnapShot, uProcess)
    Do While R
        If InStr(UCase(uProcess.szExeFile), UCase(EXEName)) <> 0 Then
            GetProcessIDByEXEName = uProcess.th32ProcessID
            Call CloseHandle(hSnapShot)
            Exit Function
        End If
        R = ProcessNext(hSnapShot, uProcess)
    Loop
    Call CloseHandle(hSnapShot)
End Function

Function ProcessTerminate(Optional lProcessID As Long, Optional lHwndWindow As Long) As Boolean
    Dim lhwndProcess As Long
    Dim lExitCode As Long
    Dim lRetVal As Long
    Dim lhThisProc As Long
    Dim lhTokenHandle As Long
    Dim tLuid As LUID
    Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINATE = &H1
    Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
    Const SE_PRIVILEGE_ENABLED = &H2

    On Error Resume Next
    If lHwndWindow Then
        'Get the process ID from the window handle
        lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
    End If
    
    If lProcessID Then
        'Give Kill permissions to this process
        lhThisProc = GetCurrentProcess
        
        OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
        LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
        'Set the number of privileges to be change
        tTokenPriv.PrivilegeCount = 1
        tTokenPriv.TheLuid = tLuid
        tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        'Enable the kill privilege in the access token of this process
        AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

        'Open the process to kill
        lhwndProcess = OpenProcess(PROCESS_TERMINATE, 0, lProcessID)
    
        If lhwndProcess Then
            'Obtained process handle, kill the process
            ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
            Call CloseHandle(lhwndProcess)
        End If
    End If
    On Error GoTo 0
End Function
