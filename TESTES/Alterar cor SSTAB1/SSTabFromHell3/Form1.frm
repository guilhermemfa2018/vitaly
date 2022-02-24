VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "SSTab - FROM HELL"
   ClientHeight    =   6135
   ClientLeft      =   3435
   ClientTop       =   2850
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9315
   Begin TabDlg.SSTab SSTab2 
      Height          =   1755
      Left            =   180
      TabIndex        =   17
      Top             =   4200
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   3096
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5475
      Left            =   1860
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   15
         Top             =   900
         Width           =   615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cut"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -120
            TabIndex        =   16
            Top             =   60
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   13
         Top             =   900
         Width           =   615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cut"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -120
            TabIndex        =   14
            Top             =   60
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   4440
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   11
         Top             =   900
         Width           =   615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cut"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   -60
            TabIndex        =   12
            Top             =   60
            Width           =   915
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4335
         Left            =   420
         TabIndex        =   1
         Top             =   780
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   7646
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "   Wider   "
         TabPicture(0)   =   "Form1.frx":001C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(1)=   "Text1"
         Tab(0).Control(2)=   "Command1"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "   Wider    "
         TabPicture(1)   =   "Form1.frx":687E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "List1"
         Tab(1).Control(1)=   "Text2"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "   Wider  "
         TabPicture(2)   =   "Form1.frx":D0E0
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Check1"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Option3"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Option2"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Option1"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).ControlCount=   4
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   885
            Left            =   -74130
            TabIndex        =   9
            Top             =   1920
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Height          =   435
            Left            =   -74100
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   3000
            Width           =   2385
         End
         Begin VB.ListBox List1 
            Height          =   645
            ItemData        =   "Form1.frx":13942
            Left            =   -74070
            List            =   "Form1.frx":13958
            TabIndex        =   7
            Top             =   1380
            Width           =   3045
         End
         Begin VB.TextBox Text2 
            Height          =   345
            Left            =   -74130
            TabIndex        =   6
            Text            =   "Text2"
            Top             =   2400
            Width           =   3135
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   195
            Left            =   270
            TabIndex        =   5
            Top             =   1530
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Left            =   270
            TabIndex        =   4
            Top             =   1890
            Width           =   3165
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   495
            Left            =   270
            TabIndex        =   3
            Top             =   2220
            Width           =   2505
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Capture WM_CTLCOLORSTATIC"
            Height          =   255
            Left            =   270
            TabIndex        =   2
            Top             =   1050
            Width           =   3045
         End
         Begin VB.Label Label1 
            Caption         =   "Test to draw on a SSTab control... ???"
            Height          =   195
            Left            =   -74520
            TabIndex        =   10
            Top             =   1560
            Width           =   2865
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   120
      Picture         =   "Form1.frx":1398C
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   285
   End
   Begin VB.Image ImageX 
      Height          =   1440
      Left            =   30
      Picture         =   "Form1.frx":14E9E
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
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

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
'    mBrush = CreatePatternBrush(Image1.Picture.Handle)
'    bgWid = Me.ScaleX(Image1.Picture.Width, vbHimetric, vbPixels)
'    bgHgt = Me.ScaleY(Image1.Picture.Height, vbHimetric, vbPixels)
    
    mBrush = CreateSolidBrush(&HD8E9EC)
    bgWid = Me.ScaleX(Image1.Picture.Width, vbHimetric, vbPixels)
    bgHgt = Me.ScaleY(Image1.Picture.Height, vbHimetric, vbPixels)

    ' Start the subclassing
    oldSSTabProc = SetWindowLong(SSTab1.Hwnd, GWL_WNDPROC, AddressOf SSTabProc)
    oldSSTabProc = SetWindowLong(SSTab2.Hwnd, GWL_WNDPROC, AddressOf SSTabProc)
End Sub

'Private Sub Form_Resize()
'    SSTab1.Move SSTab1.Left, SSTab1.Top, Me.ScaleWidth - SSTab1.Left * 2, Me.ScaleHeight - SSTab1.Top * 2
'End Sub


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

Private Sub Label2_Click(Index As Integer)
    SSTab1.Tab = Index
End Sub

Private Sub Picture1_Click(Index As Integer)
    SSTab1.Tab = Index
End Sub
