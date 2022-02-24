VERSION 5.00
Begin VB.UserControl VistaProgress 
   BackColor       =   &H00C0C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   3015
   ScaleWidth      =   7620
   ToolboxBitmap   =   "VistaProgress.ctx":0000
   Begin VB.Timer TmrAnimate 
      Interval        =   1
      Left            =   2160
      Top             =   2160
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4440
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image ImgBarBackRight 
      Height          =   225
      Left            =   3600
      Picture         =   "VistaProgress.ctx":0312
      Top             =   10
      Width           =   315
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   360
      Picture         =   "VistaProgress.ctx":2496
      Top             =   1200
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image ImgBase 
      Height          =   315
      Left            =   15
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2250
   End
   Begin VB.Image ImgBar 
      Height          =   225
      Left            =   480
      Picture         =   "VistaProgress.ctx":4549
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   4560
      Picture         =   "VistaProgress.ctx":69EA
      Top             =   1320
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image ImgCrystal 
      Height          =   195
      Left            =   360
      Picture         =   "VistaProgress.ctx":8A9D
      Top             =   1920
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image ImgBarBackMid 
      Height          =   225
      Left            =   240
      Picture         =   "VistaProgress.ctx":B063
      Stretch         =   -1  'True
      Top             =   15
      Width           =   3435
   End
   Begin VB.Image ImgBarBackLeft 
      Height          =   225
      Left            =   0
      Picture         =   "VistaProgress.ctx":D504
      Top             =   10
      Width           =   315
   End
   Begin VB.Image ImgBackRight 
      Height          =   225
      Left            =   3840
      Picture         =   "VistaProgress.ctx":F675
      Top             =   240
      Width           =   15
   End
End
Attribute VB_Name = "VistaProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const MaxValue = 100
Const progValue = 50

Dim enableAnim As Boolean
Dim pauseAnim As Boolean

Dim UsrContrlWidth As Long


Public Property Get Pause() As Boolean
    Pause = pauseAnim
End Property

Public Property Let Pause(ByVal new_Pause As Boolean)
    pauseAnim = new_Pause
    PropertyChanged "Pause"
    
    If pauseAnim = True Then
        TmrAnimate.Enabled = False
    Else
        TmrAnimate.Enabled = True
    End If
End Property

Public Property Get Enable() As Boolean
    Enable = enableAnim
End Property

Public Property Let Enable(ByVal new_Bool As Boolean)

    enableAnim = new_Bool
    PropertyChanged "Enable"
    
    If enableAnim = True Then
        ImgCrystal.Left = -2000
        TmrAnimate.Enabled = True
    Else
        picTemp.Cls
        picTemp.PaintPicture Image4.Picture, 0, 10, 30
        picTemp.PaintPicture ImgBar.Picture, 1, 10, UserControl.Width - (picTemp.Width + 50)
        ImgBase.Picture = picTemp.Image
        ImgCrystal.Left = -2000
    End If
    
End Property

Private Sub UserControl_InitProperties()
On Error Resume Next

    If UserControl.Width < 350 Then
        UserControl.Width = 350
    End If

    pauseAnim = False
    enableAnim = False

    ImgBase.Width = UserControl.Width
    picTemp.Width = UserControl.Width / MaxValue * progValue
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

    pauseAnim = PropBag.ReadProperty("Pause", False)
    enableAnim = PropBag.ReadProperty("Enable", False)

    ImgBase.Width = UserControl.Width - 20
    picTemp.Width = UserControl.Width / MaxValue * progValue
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("Pause", enableAnim, False)
    Call PropBag.WriteProperty("Enable", enableAnim, False)

    ImgBase.Width = UserControl.Width - 20
    picTemp.Width = UserControl.Width / MaxValue * progValue
End Sub

Private Sub usercontrol_Resize()
On Error Resume Next

    If UserControl.Width < 650 Then
        UserControl.Width = 650
    End If
    
    UsrContrlWidth = UserControl.Width - 75
    ImgBarBackMid.Width = UsrContrlWidth - 500
    ImgBarBackRight.Left = ImgBarBackLeft.Width + ImgBarBackMid.Width - 75
    
    ImgBase.Width = UserControl.Width - 20
    picTemp.Width = UserControl.Width / MaxValue * progValue
    
    UserControl.Height = 255

End Sub

Private Sub TmrAnimate_Timer()

    If enableAnim = True Then
    
        ImgBarBackRight.Visible = False
        TmrAnimate.Interval = 1
        
        picTemp.Cls
        ImgBase.Width = UserControl.Width - 50
        ImgBackRight.Move ImgBase.Width + 15, 10
        picTemp.PaintPicture Image4.Picture, 0, 10, 30
        picTemp.PaintPicture ImgBar.Picture, 1, 10, UserControl.Width - (picTemp.Width + 50)
        picTemp.PaintPicture ImgCrystal.Picture, ImgCrystal.Left, 25
        ImgCrystal.Move ImgCrystal.Left + 20
        
        If ImgCrystal.Left > UserControl.Width Then
            TmrAnimate.Interval = 500
            ImgCrystal.Move 0 - 1800
        End If
        
        ImgBase.Picture = picTemp.Image
        
    Else
        ImgBarBackRight.Visible = True
        TmrAnimate.Enabled = False
    End If

End Sub

