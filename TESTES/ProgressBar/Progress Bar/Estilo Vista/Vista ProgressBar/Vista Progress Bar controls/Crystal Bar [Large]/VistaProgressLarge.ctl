VERSION 5.00
Begin VB.UserControl VistaProgressLarge 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   3015
   ScaleWidth      =   7620
   ToolboxBitmap   =   "VistaProgressLarge.ctx":0000
   Begin VB.Timer TmrAnimate 
      Interval        =   1
      Left            =   2160
      Top             =   2160
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image ImgBackRight 
      Height          =   375
      Left            =   5400
      Picture         =   "VistaProgressLarge.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
   Begin VB.Image ImgBarLeft 
      Height          =   345
      Left            =   240
      Picture         =   "VistaProgressLarge.ctx":05CC
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ImgBase 
      Height          =   340
      Left            =   10
      Stretch         =   -1  'True
      Top             =   10
      Width           =   5400
   End
   Begin VB.Image ImgBackLeft 
      Height          =   375
      Left            =   0
      Picture         =   "VistaProgressLarge.ctx":08C3
      Top             =   0
      Width           =   30
   End
   Begin VB.Image ImgBack 
      Height          =   375
      Left            =   0
      Picture         =   "VistaProgressLarge.ctx":0B7C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
   Begin VB.Image ImgBarMid 
      Height          =   345
      Left            =   600
      Picture         =   "VistaProgressLarge.ctx":0E17
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image ImgBarRight 
      Height          =   345
      Left            =   1920
      Picture         =   "VistaProgressLarge.ctx":10C0
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image ImgCrystal 
      Height          =   345
      Left            =   600
      Picture         =   "VistaProgressLarge.ctx":13C1
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "VistaProgressLarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const defaultMax = 100
Const defaultValue = 0

Dim maxVal As Long
Dim ProgValue As Long


Public Property Get Max() As Long
On Error Resume Next

    Max = maxVal
    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue
End Property

Public Property Let Max(ByVal NewMax As Long)
On Error Resume Next

    maxVal = NewMax
    PropertyChanged "Max"
    
    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue
End Property

Public Property Get Value() As Long
On Error Resume Next
    Value = ProgValue

    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue
End Property

Public Property Let Value(ByVal NewValue As Long)
On Error Resume Next
    ProgValue = NewValue
    PropertyChanged "Value"
    
    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue

    If NewValue = maxVal * 0.01 Then
        TmrAnimate.Enabled = False
        ImgBase.Picture = temp.Picture
        ImgBase.Width = picTemp.Width
    ElseIf NewValue = 0 Then
        ImgBase.Visible = False
    ElseIf NewValue > maxVal Then
        ProgValue = maxVal
        TmrAnimate.Enabled = True
    Else
        TmrAnimate.Enabled = True
    End If
End Property

Private Sub UserControl_InitProperties()
On Error Resume Next
    maxVal = defaultMax
    ProgValue = defaultValue

    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

    maxVal = PropBag.ReadProperty("Max", defaultMax)
    ProgValue = PropBag.ReadProperty("Value", defaultValue)

    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("Max", maxVal, defaultMax)
    Call PropBag.WriteProperty("Value", ProgValue, defaultValue)

    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue
End Sub

Private Sub usercontrol_Resize()
On Error Resume Next

    ImgBack.Width = UserControl.Width - 30
    ImgBackRight.Left = ImgBack.Left + ImgBack.Width
    
    ImgBase.Width = UserControl.Width / maxVal * ProgValue - 15
    picTemp.Width = UserControl.Width / maxVal * ProgValue

    UserControl.Height = 375
End Sub

Private Sub TmrAnimate_Timer()

If ProgValue <> 0 Then

    TmrAnimate.Interval = 1
    picTemp.Cls
    picTemp.PaintPicture ImgBarMid.Picture, 10, 10, ImgBack.Width + ImgBack.Left
    picTemp.PaintPicture ImgBarLeft.Picture, 0, 10
    picTemp.PaintPicture ImgBarRight.Picture, UserControl.Width / maxVal * ProgValue - ImgBarRight.Width, 10
    picTemp.PaintPicture ImgCrystal.Picture, ImgCrystal.Left, 10
    ImgCrystal.Move ImgCrystal.Left + 100
    
    If ImgCrystal.Left > UserControl.Width Then
        TmrAnimate.Interval = 500
        ImgCrystal.Move 0 - (ImgCrystal.Width * 3)
    End If
    
    ImgBase.Picture = picTemp.Image
    ImgBase.Visible = True
    
ElseIf ProgValue = 0 Then
    TmrAnimate.Enabled = False
End If

End Sub

