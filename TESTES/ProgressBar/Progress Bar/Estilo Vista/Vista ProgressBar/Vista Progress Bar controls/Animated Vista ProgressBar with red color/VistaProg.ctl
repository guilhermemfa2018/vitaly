VERSION 5.00
Begin VB.UserControl VistaProg 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ScaleHeight     =   3600
   ScaleWidth      =   5880
   ToolboxBitmap   =   "VistaProg.ctx":0000
   Begin VB.Timer TmrMaduAnimateMinus 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2595
      Top             =   1350
   End
   Begin VB.Timer TmrMaduAnimatePlus 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1755
      Top             =   855
   End
   Begin VB.Image BarMainTmpRd 
      Height          =   255
      Left            =   3360
      Picture         =   "VistaProg.ctx":0312
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   15
   End
   Begin VB.Image BarMainTmpGr 
      Height          =   225
      Left            =   3240
      Picture         =   "VistaProg.ctx":23BC
      Top             =   2640
      Width           =   30
   End
   Begin VB.Image BarLeft 
      Height          =   225
      Left            =   0
      Picture         =   "VistaProg.ctx":2476
      Top             =   0
      Width           =   30
   End
   Begin VB.Image Barright 
      Height          =   225
      Left            =   2000
      Picture         =   "VistaProg.ctx":2530
      Top             =   0
      Width           =   30
   End
   Begin VB.Image RedProgRight 
      Height          =   225
      Left            =   1200
      Picture         =   "VistaProg.ctx":25EA
      Top             =   2400
      Width           =   30
   End
   Begin VB.Image RedProgLeft 
      Height          =   225
      Left            =   960
      Picture         =   "VistaProg.ctx":46D3
      Top             =   2400
      Width           =   30
   End
   Begin VB.Image BarMainRd 
      Height          =   255
      Left            =   0
      Picture         =   "VistaProg.ctx":679C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image BarmainGr 
      Height          =   225
      Left            =   0
      Picture         =   "VistaProg.ctx":8846
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image righton 
      Height          =   225
      Left            =   765
      Picture         =   "VistaProg.ctx":8900
      Stretch         =   -1  'True
      Top             =   1245
      Width           =   30
   End
   Begin VB.Image rightoff 
      Height          =   225
      Left            =   765
      Picture         =   "VistaProg.ctx":89BA
      Top             =   960
      Width           =   30
   End
   Begin VB.Image lefton 
      Height          =   225
      Left            =   555
      Picture         =   "VistaProg.ctx":8A74
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   30
   End
   Begin VB.Image leftoff 
      Height          =   225
      Left            =   540
      Picture         =   "VistaProg.ctx":8B2E
      Top             =   960
      Width           =   30
   End
   Begin VB.Image Barback 
      Height          =   225
      Left            =   15
      Picture         =   "VistaProg.ctx":8BE8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "VistaProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private barMin As Long
Private barValue As Long
Private barMax As Long
Private redProg As Boolean
Private valFormatted As Long
Private barValFormatted As Long

Private animValue As Long


Private Sub Image1_Click()

End Sub

Private Sub TmrMaduAnimateMinus_Timer()

    If Not barValue <= animValue Then
        If animValue < 20 Then
        barValue = barValue - 1
        Bar_Draw
        PropertyChanged "Value"
        Exit Sub
        End If
        barValue = barValue - FormatNumber((animValue / 20), 0)
        Bar_Draw
        PropertyChanged "Value"
    Else
        TmrMaduAnimateMinus.Enabled = False
        barValue = animValue
        Bar_Draw
        PropertyChanged "Value"
    End If
End Sub

Private Sub TmrMaduAnimatePlus_Timer()

If Not barValue >= animValue Then
    If animValue < 20 Then
    barValue = barValue + 1
    Bar_Draw
    PropertyChanged "Value"
    Exit Sub
    End If
    barValue = barValue + FormatNumber((animValue / 20), 0)
    Bar_Draw
    PropertyChanged "Value"
Else
    TmrMaduAnimatePlus.Enabled = False
    barValue = animValue
    Bar_Draw
    PropertyChanged "Value"
End If

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
    .Height = 225
    Barright.Left = .ScaleWidth - Barright.Width
    Barback.Width = .ScaleWidth
    End With
    Bar_Draw
End Sub

Public Property Let Value(ByVal val As Long)
    If val > barMax Then val = barMax
    If val < barMin Then val = barMin

        valFormatted = (val / barMax) * 100
        barValFormatted = (barValue / barMax) * 100

    If valFormatted > barValFormatted Then
        
        If valFormatted >= 80 Then
            animValue = val
            redProg = True
            
            BarmainGr.Picture = BarMainRd.Picture
            Barright.Picture = RedProgRight.Picture
            BarLeft.Picture = RedProgLeft.Picture

            TmrMaduAnimatePlus.Enabled = True
            TmrMaduAnimateMinus.Enabled = False
        Else
            animValue = val
            redProg = False
            
            BarmainGr.Picture = BarMainTmpGr.Picture
            
            TmrMaduAnimatePlus.Enabled = True
            TmrMaduAnimateMinus.Enabled = False
        End If
        
        
    ElseIf valFormatted < barValFormatted Then
    
        If valFormatted >= 80 Then
            animValue = val
            redProg = True
            
            BarmainGr.Picture = BarMainRd.Picture
            Barright.Picture = RedProgRight.Picture
            BarLeft.Picture = RedProgLeft.Picture

            TmrMaduAnimatePlus.Enabled = False
            TmrMaduAnimateMinus.Enabled = True
        Else
            animValue = val
            redProg = False
    
            BarmainGr.Picture = BarMainTmpGr.Picture
            
            TmrMaduAnimatePlus.Enabled = False
            TmrMaduAnimateMinus.Enabled = True
        End If

    Else

    barValue = val
    Bar_Draw
    PropertyChanged "Value"
    End If
'________________________________________________________________
End Property

Public Property Get Value() As Long
    Value = barValue
End Property

Public Property Let Max(ByVal val As Long)
    If val < 1 Then val = 1
    If val <= barMin Then val = barMin + 1
    barMax = val
    If Value > barMax Then Value = barMax
    Bar_Draw
    PropertyChanged "Max"
End Property
Public Property Get Max() As Long
    Max = barMax
End Property

Public Property Let Min(ByVal val As Long)
    If val >= barMax Then val = Max - 1
    If val < 0 Then val = 0
    barMin = val
    If Value < barMin Then Value = barMin
    Bar_Draw
    PropertyChanged "Min"
End Property
Public Property Get Min() As Long
    Min = barMin
End Property

Private Sub UserControl_InitProperties()
    Max = 100
    Min = 0
    Value = 50
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Max = PropBag.ReadProperty("Max", 100)
    Min = PropBag.ReadProperty("Min", 0)
    Value = PropBag.ReadProperty("Value", 50)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Max", Max, 100
    PropBag.WriteProperty "Min", Min, 0
    PropBag.WriteProperty "Value", Value, 50
End Sub

Private Sub Bar_Draw()
Dim i, s, z, y, q As Long
    i = barMax: s = barValue: z = barMax
    y = (s * 100 / z)
    q = (y * UserControl.Width / 100)
If s = 0 Then BarmainGr.Width = 15: Barright.Picture = rightoff.Picture: BarLeft.Picture = leftoff.Picture

If s >= 1 Then
    If redProg = True Then
        BarLeft.Picture = RedProgLeft.Picture
    Else
        BarLeft.Picture = lefton.Picture
    End If
BarmainGr.Width = q
End If

If s = z Then Barright.Picture = RedProgRight.Picture Else If s < z Then Barright.Picture = rightoff.Picture
End Sub
