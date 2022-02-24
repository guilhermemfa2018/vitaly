VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmSkins 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SKINS"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "FrmSkins.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6000
      OleObjectBlob   =   "FrmSkins.frx":1ADA
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   7095
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Master System"
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6480
         Top             =   2040
      End
      Begin VB.FileListBox File1 
         Height          =   4185
         Left            =   120
         Pattern         =   "*.skn"
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdAplicarSkin 
         Caption         =   "Aplicar Skin (F9)"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   1
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair (Esc)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   2
         Top             =   3600
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   3240
         OleObjectBlob   =   "FrmSkins.frx":1D0E
         TabIndex        =   4
         Top             =   360
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmSkins.frx":1D74
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmSkins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAplicarSkin_Click()
If Text1 = "Nenhum" Then
   mobjMsg.Abrir "Escolha um Skin..", , informacao, "Master System"
    Exit Sub
End If

'Copia o novo skin e salva com o nome de MySkin
FileCopy (App.Path & "\Skins\" & File1.FileName), App.Path & "\MySkin.skn"
AplicarSkin Principal, Principal.Skin1
cmdSair.SetFocus


    'Salva o nome do Skin atual
    WriteProfile "Skin", "SkinAtual", Text1, App.Path & "\CONFIG.INI"

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub File1_Click()
    Skin1.LoadSkin App.Path & "\skins\" & File1.FileName
    Skin1.ApplySkin Me.hwnd
    Text1 = File1.FileName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then 'F9
    cmdAplicarSkin_Click
End If
End Sub

Private Sub Form_Load()
AplicarSkin Me, Principal.Skin1
File1.Path = App.Path & "\skins"
Text1 = "Nenhum"

On Error Resume Next

Option1 = True
Check1 = 1
'Recupera o nome do Skin atual
Dim P_Buffer As String
Dim P_M() As String

P_Buffer = GetProfileSection("Skin", App.Path & "\CONFIG.INI")
P_M = Split(P_Buffer, vbNullChar)
SkinLabel2.Caption = Join(P_M, vbCrLf)
  
End Sub
