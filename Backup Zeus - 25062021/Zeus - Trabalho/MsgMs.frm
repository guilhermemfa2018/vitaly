VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MsgMs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "####"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "MsgMs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5490
      Begin ZEUS.Caption pergunta 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "Caption1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
         CaptionPicture  =   "MsgMs.frx":000C
      End
      Begin ZEUS.Caption informacao 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "Caption1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
         CaptionPicture  =   "MsgMs.frx":1AF6
      End
      Begin ZEUS.Caption Label1 
         Height          =   240
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         Caption         =   "####"
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
         CaptionINColor  =   0
         CaptionOutlined =   0   'False
         CaptionShadowed =   0   'False
      End
      Begin VB.CommandButton Command3 
         Caption         =   "####"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "####"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "####"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "MsgMs.frx":35E0
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin ZEUS.Caption exclamacao 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "Caption1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
         CaptionPicture  =   "MsgMs.frx":3646
      End
      Begin ZEUS.Caption critica 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "Caption1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483638
         CaptionPicture  =   "MsgMs.frx":5130
      End
   End
End
Attribute VB_Name = "MsgMs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CND As Boolean

Private Sub Command1_Click()
    Tp = "1"
    CND = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Tp = "2"
    CND = True
    Unload Me
End Sub

Private Sub Command3_Click()
    Tp = "3"
    CND = True
    Unload Me
End Sub

Private Sub Form_Activate()
    'Verifica se vai ou não utilizar skin
    'If Onde = Empty Then
    '    Exit Sub
    'Else
    'End If
AplicarSkin Me, Principal.Skin1
End Sub

Public Sub Forma()
    SkinLabel1.Caption = Label1.Caption
    SkinLabel1.Width = Label1.Width
    SkinLabel1.Height = Label1.Height
    'Calcula o Tamanho do form e frame de acordo com a mensagen
    'Se mensagen for menor que o tamanho dos tres botões não muda o tamanho
    If Label1.Width > 4575 Then
        Me.Width = Label1.Left + Label1.Width + 480
        Frame1.Width = Me.Width - 105
    End If
    'Calcula a altura do form e frame de acordo com a mensagen e
    'Posição dos butões "Top"
    If Label1.Height > 615 Then
        Command1.Top = Label1.Top + Label1.Height + 100
        Command2.Top = Label1.Top + Label1.Height + 100
        Command3.Top = Label1.Top + Label1.Height + 100
        Frame1.Height = Me.Height + 100
        Me.Height = Command1.Top + Command1.Height + 480
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Se CND = False então cancela o fechamento para que resposta não seja em branco
    If CND = False Then Cancel = 1
End Sub
