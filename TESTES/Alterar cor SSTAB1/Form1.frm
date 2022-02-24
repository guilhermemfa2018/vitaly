VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Form1 
   Caption         =   "SSTab - FROM HELL"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17025
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   17025
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).Control(1)=   "List1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Option1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Option2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Option3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Check1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CheckBox Check1 
         Caption         =   "Capture WM_CTLCOLORSTATIC"
         Height          =   255
         Left            =   570
         TabIndex        =   9
         Top             =   570
         Width           =   3045
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   495
         Left            =   570
         TabIndex        =   7
         Top             =   1740
         Width           =   2505
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   570
         TabIndex        =   6
         Top             =   1410
         Width           =   3165
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   570
         TabIndex        =   5
         Top             =   1050
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   -74490
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1740
         Width           =   3135
      End
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "Form1.frx":0054
         Left            =   -74430
         List            =   "Form1.frx":006A
         TabIndex        =   3
         Top             =   720
         Width           =   3045
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   -74100
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1920
         Width           =   2385
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   885
         Left            =   -74130
         TabIndex        =   1
         Top             =   840
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "Test to draw on a SSTab control... ???"
         Height          =   195
         Left            =   -74400
         TabIndex        =   8
         Top             =   450
         Width           =   3285
      End
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   11040
      Picture         =   "Form1.frx":009E
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   30
      Picture         =   "Form1.frx":0395
      Top             =   30
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************************
'  Custom SSTab control Demo
'  -------------------------
'
'     Written by: Garrett Sever (aka "The Hand")
'           Date: 8/19/01
'       Revision: 4/24/03 - Boredom set in. Revised it a little and eliminated
'                           the flicker completely. It still lags a bit when its
'                           resized really big, but that's the danger of bitblt
'                           color replacements over large areas.
'
' *********************************************************************************
'     Feel free to use this source code as you wish in your projects, however
'     if you publish it, either on a website, forum, book, etc. give credit where
'     its due.
' *********************************************************************************

Option Explicit

' *********************************************************************************
'  API Declarations...
' *********************************************************************************
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal Hwnd As Long, ByVal lpRect As Long) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Const GWL_WNDPROC = (-4)

' *********************************************************************************
'  Module level variables...
' *********************************************************************************
Private bgWid As Long
Private bgHgt As Long
Private oldSSTabProc As Long
Private mBrush As Long

Private Sub Form_Load()
    ' grab our background image's dimensions for later use
    mBrush = CreatePatternBrush(Image2.Picture.Handle)
    bgWid = Me.ScaleX(Image2.Picture.Width, vbHimetric, vbPixels)
    bgHgt = Me.ScaleY(Image2.Picture.Height, vbHimetric, vbPixels)
    
    ' Start the subclassing
    oldSSTabProc = SetWindowLong(SSTab1.Hwnd, GWL_WNDPROC, AddressOf SSTabProc)
End Sub

Private Sub Form_Resize()
    SSTab1.Move SSTab1.Left, SSTab1.Top, Me.ScaleWidth - SSTab1.Left * 2, Me.ScaleHeight - SSTab1.Top * 2
End Sub


Friend Function NewSSTabProc(ByVal sstHwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    Dim aRect       As RECT
    Dim updateRect  As RECT
    Dim destDC      As Long
    Dim tempDC      As Long
    Dim tempBmp     As Long
    Dim origDC      As Long
    Dim origBmp     As Long
    Dim maskDC      As Long
    Dim maskBmp     As Long
    Dim memDC       As Long
    Dim memBmp      As Long
    
    Dim wid         As Long
    Dim hgt         As Long
    Dim x           As Long
    Dim y           As Long
    Dim aControl    As Control
    
    Dim origBrush As Long
    Dim origColor As Long
    
    On Error Resume Next
    If wMsg = &HF Then  'WM_PAINT
        
        GetUpdateRect sstHwnd, updateRect, False
        With updateRect
            Debug.Print "(" & .Left & "," & .Top & ")-(" & .Right & "," & .Bottom & ")"
        End With
        
        ' get the SSTab's device context
        destDC = GetDC(sstHwnd)
        
        ' get the SSTab's window dimensions
        GetWindowRect sstHwnd, aRect
        wid = aRect.Right - aRect.Left
        hgt = aRect.Bottom - aRect.Top
        
        ' create our other temporary device contexts.
        maskDC = CreateCompatibleDC(destDC)
        maskBmp = CreateBitmap(wid, hgt, 1, 1, ByVal 0&)
        memDC = CreateCompatibleDC(destDC)
        memBmp = CreateCompatibleBitmap(destDC, wid, hgt)
        tempDC = CreateCompatibleDC(destDC)
        tempBmp = CreateCompatibleBitmap(destDC, wid, hgt)
        origDC = CreateCompatibleDC(destDC)
        origBmp = CreateCompatibleBitmap(destDC, wid, hgt)
        
        ' delete the temporary 1x1 bitmap and put our (wid)x(hgt) ones in
        DeleteObject SelectObject(maskDC, maskBmp)
        DeleteObject SelectObject(memDC, memBmp)
        DeleteObject SelectObject(tempDC, tempBmp)
        DeleteObject SelectObject(origDC, origBmp)
        
        ' Call the control's original handler... paints the control on our back buffer
        CallWindowProc oldSSTabProc, sstHwnd, wMsg, origDC, lParam

        ' This helps our mask to correctly calculate the b & w pixels of
        '  our mask. Only really works in Win98 and greater... and even then
        '  it is sometimes flakey... may need to loop thru x & y and use
        '  GetPixel/SetPixel to create mask if it is not generated properly.
        origColor = SetBkColor(destDC, GetSysColor(15))
        SetBkColor origDC, GetSysColor(15)
        ' create a b&w pixel mask - background color is white, everything else
        '  is black.
        BitBlt maskDC, 0, 0, wid, hgt, origDC, 0, 0, vbSrcCopy
                

        ' select the pattern brush into the DC and pattern blit
        origBrush = SelectObject(tempDC, mBrush)
        PatBlt tempDC, 0, 0, wid, hgt, vbPatCopy
        SelectObject tempDC, origBrush
        
        ' clean up our original image of the control so only the non background
        '  color parts are showing... make everything else white.
        BitBlt memDC, 0, 0, wid, hgt, maskDC, 0, 0, vbSrcCopy
        BitBlt memDC, 0, 0, wid, hgt, origDC, 0, 0, vbSrcPaint
        

        'punch the hole for our control image
        BitBlt tempDC, 0, 0, wid, hgt, maskDC, 0, 0, vbMergePaint
        'put the control images back in
        BitBlt tempDC, 0, 0, wid, hgt, memDC, 0, 0, vbSrcAnd
        'copy our new version back to the control
        BitBlt destDC, 0, 0, wid, hgt, tempDC, 0, 0, vbSrcCopy

        ' clean up all of our used graphical resources (VERY IMPORTANT!!!)
        DeleteDC tempDC
        DeleteObject tempBmp
        DeleteDC maskDC
        DeleteObject maskBmp
        DeleteDC memDC
        DeleteObject memBmp
        DeleteDC origDC
        DeleteObject origBmp
        
        ' Replace the original background color
        SetBkColor destDC, origColor
        ' Release the SSTab's device context back to the system
        ReleaseDC sstHwnd, destDC
        
        ValidateRect sstHwnd, 0
                
        On Error GoTo 0
    ElseIf wMsg = &H2 Then 'WM_DESTROY
        DeleteObject mBrush
        SetWindowLong sstHwnd, GWL_WNDPROC, oldSSTabProc
        NewSSTabProc = CallWindowProc(oldSSTabProc, sstHwnd, wMsg, wParam, lParam)
    ElseIf wMsg = &H138 And _
           Check1.Value Then    '&H138 = WM_CTLCOLORSTATIC
        SetBkMode wParam, 1     ' make the text draw transparent
        NewSSTabProc = mBrush   ' return the background brush
    Else
        NewSSTabProc = CallWindowProc(oldSSTabProc, sstHwnd, wMsg, wParam, lParam)
    End If
    On Error GoTo 0
End Function

