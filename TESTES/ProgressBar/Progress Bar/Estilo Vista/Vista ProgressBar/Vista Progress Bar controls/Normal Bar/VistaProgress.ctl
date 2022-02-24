VERSION 5.00
Begin VB.UserControl VistaProgress 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   3015
   ScaleWidth      =   7620
   ToolboxBitmap   =   "VistaProgress.ctx":0000
   Begin VB.Timer tmrAnimate 
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
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image9 
      Height          =   195
      Left            =   5640
      Picture         =   "VistaProgress.ctx":0312
      Top             =   1920
      Width           =   285
   End
   Begin VB.Image ImgLine 
      Height          =   15
      Left            =   4080
      Picture         =   "VistaProgress.ctx":2481
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image ImgBackRight 
      Height          =   225
      Left            =   5400
      Picture         =   "VistaProgress.ctx":44CE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image ImgBarLeft 
      Height          =   195
      Left            =   240
      Picture         =   "VistaProgress.ctx":6554
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image ImgBase 
      Height          =   315
      Left            =   15
      Stretch         =   -1  'True
      Top             =   15
      Width           =   5400
   End
   Begin VB.Image ImgBackLeft 
      Height          =   225
      Left            =   0
      Picture         =   "VistaProgress.ctx":86CC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
   Begin VB.Image ImgBack 
      Height          =   225
      Left            =   0
      Picture         =   "VistaProgress.ctx":A77F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3435
   End
   Begin VB.Image ImgBarMid 
      Height          =   225
      Left            =   600
      Picture         =   "VistaProgress.ctx":CC20
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image ImgBarRight 
      Height          =   195
      Left            =   1920
      Picture         =   "VistaProgress.ctx":EDC4
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "VistaProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const defaultMax = 100
Const defaultValue = 100

Dim MaxValue As Long
Dim ProgValue As Long


Public Property Get Max() As Long
On Error Resume Next

    Max = MaxValue

    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue
End Property

Public Property Let Max(ByVal NewMax As Long)
On Error Resume Next

    MaxValue = NewMax
    PropertyChanged "Max"

    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue
End Property

Public Property Get Value() As Long
On Error Resume Next

    Value = ProgValue
    
    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue
End Property

Public Property Let Value(ByVal NewValue As Long)
On Error Resume Next

    ProgValue = NewValue
    PropertyChanged "Value"
    
    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue

    If NewValue = 0 Then
        tmrAnimate.Enabled = False
        ImgBase.Visible = False
    ElseIf NewValue > MaxValue Then
        ProgValue = MaxValue
        tmrAnimate.Enabled = True
    ElseIf NewValue = MaxValue * 0.01 Then
        tmrAnimate.Enabled = False
        ImgBase.Picture = Temp.Picture
    Else
        tmrAnimate.Enabled = True
    End If

End Property

Private Sub UserControl_InitProperties()
On Error Resume Next

    MaxValue = defaultMax
    ProgValue = defaultValue
    
    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

    MaxValue = PropBag.ReadProperty("Max", defaultMax)
    ProgValue = PropBag.ReadProperty("Value", defaultValue)

    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("Max", MaxValue, defaultMax)
    Call PropBag.WriteProperty("Value", ProgValue, defaultValue)

    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue
End Sub

Private Sub usercontrol_Resize()
On Error Resume Next

    imgBack.Width = UserControl.Width - 15
    imgBackRight.Left = UserControl.Width - 15

    ImgBase.Width = UserControl.Width / MaxValue * ProgValue - 15
    picTemp.Width = UserControl.Width / MaxValue * ProgValue

    UserControl.Height = 230
End Sub

Private Sub tmrAnimate_Timer()

    If ProgValue <> 0 Then
    
        tmrAnimate.Interval = 1
        picTemp.Cls
        picTemp.PaintPicture ImgBarMid.Picture, 10, 10, imgBack.Width + imgBack.Left
        picTemp.PaintPicture ImgBarLeft.Picture, 0, 10
        picTemp.PaintPicture ImgBarRight.Picture, UserControl.Width / MaxValue * ProgValue - ImgBarRight.Width, 10
        picTemp.PaintPicture ImgLine.Picture, 0, 200, UserControl.Width
        picTemp.PaintPicture ImgLine.Picture, 10, 200, UserControl.Width
        ImgBase.Picture = picTemp.Image
        ImgBase.Visible = True
        
    ElseIf ProgValue = 0 Then
        tmrAnimate.Enabled = False
    End If

End Sub

