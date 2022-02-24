VERSION 5.00
Begin VB.Form AtualizaZEUS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2040
   End
End
Attribute VB_Name = "AtualizaZEUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private camZEUSso As String

Private Sub atualizaEXE()
    'Fecha ZEUS
    Shell "taskkill /F /im ZEUS.exe"
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    
    Dim shell1, strOS, strVerKey, strVersion
    Set shell1 = CreateObject("WScript.Shell")
    strOS = shell1.ExpandEnvironmentStrings("%OS%")
    If strOS = "Windows_NT" Then
        strVerKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
        strVersion = shell1.regread(strVerKey & "ProductName")
    Else
        strVerKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\"
        strVersion = shell1.regread(strVerKey & "ProductName")
    End If
    Set shell1 = Nothing
    
    
    camZEUSso = Reg.regread("HKEY_LOCAL_MACHINE\Software\ZEUS\sPathZEUS")
    
    atualizaEXE
End Sub

Private Sub Timer1_Timer()
    ' Copia Arquivos e se existir no destino sobre-escreve

    ' Criando a Variavel - File System Object e Drive
    ' Para Controle de Discos (HDs) e Arquivos no Sistema
    
    Dim strOrigem As String
    Dim strDestino As String
    strOrigem = App.Path & "\ZEUS.exe"
    strDestino = camZEUSso
    
    Dim fso As New FileSystemObject
    Dim drvDrive As Drive
    fso.CopyFile strOrigem, strDestino, True
    ' Abre Programa ZEUS
    Timer1.Enabled = False
    Timer2.Enabled = True
    Exit Sub
ErroCopiaARQ:
    ' Mostra ERRO
    MsgBox Err.Description & " - " & Err.Number, vbCritical
End Sub

Private Sub Timer2_Timer()
    Shell camZEUSso, vbNormalFocus
    Timer2.Enabled = False
    Unload Me
    End
End Sub
