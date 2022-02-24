VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmLocalizar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizador de Imagens"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   Icon            =   "Visualizador de Imagens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgFig1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   6375
      Left            =   3600
      ScaleHeight     =   6345
      ScaleWidth      =   7665
      TabIndex        =   5
      Top             =   360
      Width           =   7695
      Begin VB.Image imgFig 
         BorderStyle     =   1  'Fixed Single
         Height          =   6375
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7695
      End
   End
   Begin VB.CommandButton Cmd1 
      Appearance      =   0  'Flat
      Caption         =   "Confirmar (F9)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdSair 
      Appearance      =   0  'Flat
      Caption         =   "Sair (Esc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   120
      Pattern         =   "*.bmp;*.jpg;*.gif"
      TabIndex        =   2
      Top             =   3960
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Visualizador de Imagens.frx":1CFA
      TabIndex        =   7
      Top             =   3720
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Visualizador de Imagens.frx":1D82
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd1_Click()
On Error GoTo ErrHandler
    CaMinho = Servidor & ":"
    FileCopy (File1.Path & "\" & File1.FileName), App.Path & "\PlanoDeFundo.jpg"
    Img
    Unload Me
Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim X As Picture

Set X = LoadPicture(File1.Path & "\" & File1.FileName)
Set imgFig.Picture = X
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmd1_Click
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then 'F9
    Cmd1_Click
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
AplicarSkin Me, Principal.Skin1

Dir1.Path = App.Path & "\Papel de Parede"
Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

