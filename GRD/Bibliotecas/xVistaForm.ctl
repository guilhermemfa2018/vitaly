VERSION 5.00
Begin VB.UserControl xVistaForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DEDEDE&
   BackStyle       =   0  'Transparent
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   450
   ScaleWidth      =   4830
   ToolboxBitmap   =   "xVistaForm.ctx":0000
End
Attribute VB_Name = "xVistaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MAXIMIZAR JANELA
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

' 0 NORMAL
' 1 MINIMIZADO
' 2 MAXIMINIZADO

Private Type FORMRECT
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Dim FORMRECT As FORMRECT


Private Declare Function GetActiveWindow Lib "user32" () As Integer
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Enum GRADIENT_DIR1
    Horizontal = &H0
    Vertical = &H1
End Enum

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim gCTitleDir As GRADIENT_DIR1
Dim Janela_Ativa As Boolean

    '****************************************************************
    ' Project:      Creates an Ownerdrawn Vista Style Form control
    ' Programmer:   Alexander Mungall
    ' UserControl:  xVistaForm
    ' Email:        goober_mpc@hotmail.com
    '----------------------------------------------------------------
    ' xVistaForm Copyright© Alexander Mungall, All Rights Reserved
    ' Feel free to use this code for personal use in anyway you see
    ' fit, but please give credit where credit is due...
    ' It's all I ask.
    '****************************************************************
    Option Explicit
    
    Private xlFormGradientBottom As Long
    Private xlFormGradientTop As Long
    Private xlFormInnerBorder As Long
    Private xlFormMiddleBorder As Long
    Private xlFormOuterBorder As Long
    Private xlButtonGradientBottom As Long
    Private xlButtonGradientBottomClicked As Long
    Private xlButtonGradientBottomHover As Long
    
    ' Booleans
    Private bEnableCloseButton As Boolean
    Private bEnableMaximiseButton As Boolean
    Private bEnableMinimiseButton As Boolean
    Private bCloseButton As Boolean
    Private bCloseButtonClicked As Boolean
    Private bCloseButtonHover As Boolean
    Private bDisplayIcon As Boolean
    Private bFontBold As Boolean
    Private bFontItalic As Boolean
    Private bFontStrikeThru As Boolean
    Private bFontUnderline As Boolean
    Private bMaximiseButton As Boolean
    Private bMaximiseButtonClicked As Boolean
    Private bMaximiseButtonHover As Boolean
    Private bMinimiseButton As Boolean
    Private bMinimiseButtonClicked As Boolean
    Private bMinimiseButtonHover As Boolean
    Private bMouseClicked As Boolean
    Private bMouseOnForm As Boolean
    Private bPaintForm As Boolean
    Private bRightClick As Boolean
    Private bSystemTray As Boolean
    Private bTransparency As Boolean
    Private bUnloadForm As Boolean
    
    ' Controls
    Private imgFormPic As Image
    Private myForm As Form
    Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1
    Private WithEvents picForm As PictureBox
Attribute picForm.VB_VarHelpID = -1
    Private WithEvents TmrMouseMove As Timer
Attribute TmrMouseMove.VB_VarHelpID = -1
    
    ' Doubles
    Private dFontSize As Double
    
    ' Enums
    Public Enum xVista_Type
        Vista_Aero = 0
        Vista_Basic = 1
    End Enum
    
    Public Enum xVistaStyles
        VistaBlue = 0
        VistaDark = 1
        VistaCustom = 2
    End Enum
    
    Private xVisualStyles As xVistaStyles
    Private xVisual_Type As xVista_Type

    ' Integers
    Private I As Integer
    Private iHorizontal As Integer
    Private iNumControls As Integer
    Private iTransparency As Integer
    Private iVertical As Integer
    
    ' Longs
    Private Col As Long
    Private lBottomR As Long
    Private lBottomG As Long
    Private lBottomB As Long
    Private lButtonGradientBottom(7) As Long
    Private lButtonGradientBottomClicked(7) As Long
    Private lButtonGradientBottomHover(7) As Long
    Private lButtonGradientTop As Long
    Private lButtonGradientTopClicked As Long
    Private lButtonGradientTopHover As Long
    Private lButtonInnerBorder As Long
    Private lButtonOuterBorder As Long
    Private lCloseButtonGradientBottom(7) As Long
    Private lCloseButtonGradientBottomClicked(7) As Long
    Private lCloseButtonGradientBottomHover(7) As Long
    Private lCloseButtonGradientTop As Long
    Private lCloseButtonGradientTopClicked As Long
    Private lCloseButtonGradientTopHover As Long
    Private lCloseButtonInnerBorder As Long
    Private lCloseButtonOuterBorder As Long
    Private lTopR As Long
    Private lTopG As Long
    Private lTopB As Long
    Private lFormCaptionColor As Long
    Private lFormGradientBottom As Long
    Private lFormGradientTop As Long
    Private lFormInnerBorder As Long
    Private lFormMaxHeight As Long
    Private lFormMinHeight As Long
    Private lFormMaxWidth As Long
    Private lFormMinWidth As Long
    Private lFormMiddleBorder As Long
    Private lFormOuterBorder As Long
    Private lngReturnValue As Long
    Private lSysTrayMenu As Long
    
    ' Strings
    Private sFormCaption As String

    ' Move a Titleless Window
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Const HTCAPTION = 2
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const WM_SYSCOMMAND = &H112
    Private Const SC_MOVE = &HF010
    Private Const WM_POPUPSYSTEMMENU = &H313
    
    ' Make a Semi Transparent Form
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const G = (-20)
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2
    Private Const ULW_COLORKEY = &H1
    Private Const ULW_ALPHA = &H2
    Private Const ULW_OPAQUE = &H4
    Private Const WS_EX_LAYERED = &H80000
    
    ' Show a Form in the Taskbar
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_APPWINDOW = &H40000
    Private Const SW_HIDE = 0
    Private Const SW_NORMAL = 1

    ' Types
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    ' Functions
    Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

    ' Local constants and variables declarations
    Const BorderPixels = 5
    Private TwipX As Single, TwipY As Single, BorderWidth As Single, BorderHeight As Single

    ' Type passed to Shell_NotifyIcon
    Private Type NotifyIconData
      Size As Long
      Handle As Long
      ID As Long
      flags As Long
      CallBackMessage As Long
      Icon As Long
      Tip As String * 64
    End Type

    ' Region combine consts for making the Transparent areas
    Private Const RGN_AND = 1   'Combines an intersection
    Private Const RGN_OR = 2    'Creates a union of two regions
    Private Const RGN_XOR = 3   'Creations a union of two objects with the exception of overlapping
    Private Const RGN_DIFF = 4  'Combines two regions
    Private Const RGN_COPY = 5  'Copy a region

    ' Declarations for making the Transparent areas
    Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    
    'Our declarations for retrieving colors
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

    ' Constants for managing System Tray tasks
    Private Const AddIcon = &H0
    Private Const ModifyIcon = &H1
    Private Const DeleteIcon = &H2
    
    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDBLCLK = &H203
    Private Const WM_LBUTTONDOWN = &H201
    Private Const WM_LBUTTONUP = &H202
    
    Private Const WM_RBUTTONDBLCLK = &H206
    Private Const WM_RBUTTONDOWN = &H204
    Private Const WM_RBUTTONUP = &H205
    
    Private Const MessageFlag = &H1
    Private Const IconFlag = &H2
    Private Const TipFlag = &H4
    
    Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, Data As NotifyIconData) As Boolean
    
    Private Data As NotifyIconData

    ' Menu Functions
    Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpNewItem$) As Long
    Private Declare Function CreatePopupMenu Lib "user32" () As Long
    Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Long
    'Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu&, ByVal wFlags&, ByVal X&, ByVal Y&, ByVal nReserved&, ByVal Hwnd&, ByVal lpRect&) As Long
    Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu&, ByVal wIDEnableItem&, ByVal wEnable&) As Long
    Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
    Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
    Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
    Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
    
    ' Menu Constants, Enums & Events
    Private Const MFS_BYCOMMAND As Long = &H0&
    Private Const MFS_CHECKED As Long = &H8&
    Private Const MFS_DEFAULT As Long = &H1000&
    Private Const MFS_DISABLED As Long = &H2&
    Private Const MFS_ENABLED As Long = &H0
    Private Const MFS_GRAYED As Long = &H1&
    Private Const MFS_STRING As Long = &H0&
    Private Const MFS_SEPARATOR As Long = &H800&
    Private Const MIIM_CHECKMARKS As Long = &H8
    Private Const MIIM_DATA = &H20
    Private Const MIIM_ID = &H2
    Private Const MIIM_STATE As Long = &H1&
    Private Const MIIM_TYPE As Long = &H10&
    Private Const TPM_RETURNCMD As Long = &H100&
    Public Enum menuEStates
        xDisabled = 1
        xGrayed = 2
    End Enum

    Private Type MENUITEMINFO
        cbSize As Long
        fMask As Long
        fType As Long
        fState As Long
        wID As Long
        hSubMenu As Long
        hbmpChecked As Long
        hbmpUnchecked As Long
        dwItemData As Long
        dwTypeData As String
        cch As Long
    End Type

    Event Execute(ByVal ID As Long)
       
    '****************************************************************
    ' Gradient Code: Written by Mark Gordon (msg555)
    '----------------------------------------------------------------
    ' Copyright© Mark Gordon, All Rights Reserved
    '----------------------------------------------------------------
    Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
    Private Const DIB_RGB_COLORS = 0&
    Private Const BI_RGB = 0&
    
    Private Type BITMAPINFOHEADER '40 bytes
       biSize As Long
       biWidth As Long
       biHeight As Long
       biPlanes As Integer
       biBitCount As Integer
       biCompression As Long
       biSizeImage As Long
       biXPelsPerMeter As Long
       biYPelsPerMeter As Long
       biClrUsed As Long
       biClrImportant As Long
    End Type
    
    Private Type RGBQUAD
       rgbBlue As Byte
       rgbGreen As Byte
       rgbRed As Byte
       rgbReserved As Byte
    End Type
    
    Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
      bmiColors As RGBQUAD
    End Type
    
    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As Long
        xExt As Long
        yExt As Long
    End Type
    Private Type Guid
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type
    
    Private Enum Blends
        RGBBlend = 0
        HSLBlend = 1
    End Enum

Private Function CreateGradient(Width As Long, Height As Long, LeftToRight As Boolean, LeftTopColor As Long, RightBottomColor As Long, BlendType As Blends) As StdPicture
    Dim hBmp As Long, Bits() As Byte
    Dim RS As Byte, GS As Byte, BS As Byte 'Start RGB
    Dim RE As Byte, GE As Byte, BE As Byte 'End RGB
    Dim HS As Single, SS As Single, LS As Single 'Start HSL
    Dim He As Single, SE As Single, LE As Single 'End HSL
    Dim Rc As Byte, GC As Byte, BC As Byte 'Current iteration RGB
    Dim X As Long, Y As Long
    ReDim Bits(0 To 3, 0 To Width - 1, 0 To Height - 1)
    
    RgbCol LeftTopColor, RS, GS, BS
    RgbCol RightBottomColor, RE, GE, BE
    
    If BlendType = RGBBlend Then
        If LeftToRight Then
            For X = 0 To Width - 1
                Rc = (1& * RS - RE) * ((Width - 1 - X) / (Width - 1)) + RE
                GC = (1& * GS - GE) * ((Width - 1 - X) / (Width - 1)) + GE
                BC = (1& * BS - BE) * ((Width - 1 - X) / (Width - 1)) + BE
                For Y = 0 To Height - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        Else
            For Y = 0 To Height - 1
                Rc = (1& * RS - RE) * ((Height - 1 - Y) / (Height - 1)) + RE
                GC = (1& * GS - GE) * ((Height - 1 - Y) / (Height - 1)) + GE
                BC = (1& * BS - BE) * ((Height - 1 - Y) / (Height - 1)) + BE
                For X = 0 To Width - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        End If
    ElseIf BlendType = HSLBlend Then
        RGBToHSL RS, GS, BS, HS, SS, LS
        RGBToHSL RE, GE, BE, He, SE, LE
        If LeftToRight Then
            For X = 0 To Width - 1
                HSLToRGB (1& * HS - He) * ((Width - 1 - X) / (Width - 1)) + He, _
                        (1& * SS - SE) * ((Width - 1 - X) / (Width - 1)) + SE, _
                        (1& * LS - LE) * ((Width - 1 - X) / (Width - 1)) + LE, _
                        Rc, GC, BC
                For Y = 0 To Height - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        Else
            For Y = 0 To Height - 1
                HSLToRGB (1& * HS - He) * ((Height - 1 - Y) / (Height - 1)) + He, _
                        (1& * SS - SE) * ((Height - 1 - Y) / (Height - 1)) + SE, _
                        (1& * LS - LE) * ((Height - 1 - Y) / (Height - 1)) + LE, _
                        Rc, GC, BC
                For X = 0 To Width - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        End If
    End If

    Dim BI As BITMAPINFO
    With BI.bmiHeader
        .biSize = Len(BI.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    hBmp = CreateBitmap(Width, Height, 1&, 32&, ByVal 0)
    SetDIBits 0&, hBmp, 0, Abs(BI.bmiHeader.biHeight), Bits(0, 0, 0), BI, DIB_RGB_COLORS

    Dim IGuid As Guid, PicDst As PictDesc
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    OleCreatePictureIndirect PicDst, IGuid, True, CreateGradient
End Function

'Helper Functions
Private Sub RgbCol(Col As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    R = Col And &HFF&
    G = (Col And &HFF00&) \ &H100&
    B = (Col And &HFF0000) \ &H10000
End Sub

Private Sub RGBToHSL(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, H As Single, S As Single, L As Single)
    'http://www.vbAccelerator.com
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single

    rR = R / 255: rG = G / 255: rB = B / 255

    '{Given: rgb each in [0,1].
    ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    L = (Max + Min) / 2    '{This is the lightness}
    '{Next calculate saturation}
    If Max = Min Then
        'begin {Acrhomatic case}
        S = 0
        H = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
             '{First calculate the saturation.}
        If L <= 0.5 Then
            S = (Max - Min) / (Max + Min)
        Else
            S = (Max - Min) / (2 - Max - Min)
        End If
        
        '{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            H = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            H = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
        'Debug.Print h
        'h = h * 60
        'If h < 0# Then
        '     h = h + 360            '{Make degrees be nonnegative}
        'End If
    'end {Chromatic Case}
    End If
'end {RGB_to_HLS}
End Sub

Private Sub HSLToRGB(ByVal H As Single, ByVal S As Single, ByVal L As Single, R As Byte, G As Byte, B As Byte)
    'http://www.vbAccelerator.com
    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single
    
    If S = 0 Then
        ' Achromatic case:
        rR = L: rG = L: rB = L
    Else
        ' Chromatic case:
        ' delta = Max-Min
        If L <= 0.5 Then
            'S = (Max - Min) / (Max + Min)
            ' Get Min value:
            Min = L * (1 - S)
        Else
            'S = (Max - Min) / (2 - Max - Min)
            ' Get Min value:
            Min = L - S * (1 - L)
        End If
        ' Get the Max value:
        Max = 2 * L - Min
       
        ' Now depending on sector we can evaluate the h,l,s:
        If (H < 1) Then
            rR = Max
            If (H < 0) Then
                rG = Min
                rB = rG - H * (Max - Min)
            Else
                rB = Min
                rG = H * (Max - Min) + rB
            End If
        ElseIf (H < 3) Then
            rG = Max
            If (H < 2) Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
            Else
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (H < 4) Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
            Else
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
    End If
    R = rR * 255: G = rG * 255: B = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR > rG) Then
        If (rR > rB) Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If (rB > rG) Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR < rG) Then
        If (rR < rB) Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If (rB < rG) Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If
End Function

Private Sub AddIconToTray()
  Data.Size = Len(Data)
  Data.Handle = hwnd
  Data.ID = vbNull
  Data.flags = IconFlag Or TipFlag Or MessageFlag
  Data.CallBackMessage = WM_MOUSEMOVE
  Data.Icon = Icon
  Data.Tip = sFormCaption & vbNullChar
  Call Shell_NotifyIcon(AddIcon, Data)
End Sub

Private Sub DeleteIconFromTray()
  Call Shell_NotifyIcon(DeleteIcon, Data)
End Sub

Public Function MakeSemiTransparent(ByVal hwnd As Long, ByVal Perc As Integer) As Long
    Dim msg As Long
    On Error Resume Next
     
    Perc = ((100 - Perc) / 100) * 255
    If Perc < 0 Or Perc > 255 Then
        MakeSemiTransparent = 1
    Else
        msg = GetWindowLong(hwnd, G)
        msg = msg Or WS_EX_LAYERED
        SetWindowLong hwnd, G, msg
        
        ' Set the Form header bottom colour
        Col = myForm.BackColor
        lBottomR = (Col And &HFF&)
        lBottomG = (Col And &HFF00&) / &H100
        lBottomB = (Col And &HFF0000) / &H10000
        
        SetLayeredWindowAttributes hwnd, RGB(lBottomR, lBottomG, lBottomB), Perc, LWA_ALPHA
        MakeSemiTransparent = 0
    End If
    If Err Then
        MakeSemiTransparent = 2
    End If
End Function

Private Function MakeTransparent(ByRef Frm As Form, ByVal TrnsColor As Long)
    Frm.BorderStyle = 0
     
    Dim ScaleSize As Long
    Dim Width, Height As Long 'Width and height of the image on our form
    Dim rgnMain As Long 'The main region which will be skinned then will be applied to our form
    Dim X, Y As Long 'Variables containing current X, Y in loop below
    Dim rgnPixel As Long 'A single pixel to be cut out of our image
    Dim rgbColor As Long 'A variable to store a color in the loop below
    Dim dcMain As Long 'The temporary DC of where all the skinning takes place
    Dim bmpMain As Long '1x1 bitmap created when dcMain is created
    
    ScaleSize = Frm.ScaleMode
    Frm.ScaleMode = 3 'Set the scale mode to pixels
    
    'This will get the height and width of the image on our form
    Width = Frm.ScaleX(Frm.Width, vbTwips, vbPixels)
    Height = Frm.ScaleY(Frm.Height, vbTwips, vbPixels)
    'vbHimetric
'    Frm.Width = Width * Screen.TwipsPerPixelX
'    Frm.Height = Height * Screen.TwipsPerPixelY
    
    'This will create our basic region to fit the dimensions of our
    'forms image
    rgnMain = CreateRectRgn(0, 0, Width, Height)
    
    'This will create a DC where all the skinning takes place
    dcMain = CreateCompatibleDC(Frm.hdc)
    bmpMain = SelectObject(dcMain, Frm.Picture.Handle)
    
    For Y = 0 To 4
        For X = 0 To 4 'Width
            rgbColor = GetPixel(dcMain, X, Y) 'Gets the color of a pixel on dcMain
            
            If rgbColor = TrnsColor Then 'If we found a mask color then cut it out of dcMain
                rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1) 'Create a region of a single pixel
                CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR 'Cut it out
                DeleteObject rgnPixel 'Delete it from the memory
            End If
        Next X
    Next Y
    
    For Y = 0 To 4
        For X = (Width - 5) To Width 'Width
            rgbColor = GetPixel(dcMain, X, Y) 'Gets the color of a pixel on dcMain
            
            If rgbColor = TrnsColor Then 'If we found a mask color then cut it out of dcMain
                rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1) 'Create a region of a single pixel
                CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR 'Cut it out
                DeleteObject rgnPixel 'Delete it from the memory
            End If
        Next X
    Next Y
     
    'Clear up our memory
    SelectObject dcMain, bmpMain
    DeleteDC dcMain
    DeleteObject bmpMain
    
    If rgnMain <> 0 Then
        SetWindowRgn Frm.hwnd, rgnMain, True 'Apply rgnMain to our form
    End If
     
    Frm.ScaleMode = ScaleSize
End Function

Private Function RemoveTransparent(ByRef Frm As Form)

    Dim Width, Height As Long
    Dim rgnMain As Long
    
    'Get size of form
    Width = Frm.ScaleWidth
    Height = Frm.ScaleHeight
    
    rgnMain = CreateRectRgn(0, 0, Width, Height) 'Create a plain old region
    SetWindowRgn Frm.hwnd, rgnMain, True 'Apply to our window
 
End Function

Private Sub ShowInTheTaskbar(hwnd As Long, bShow As Boolean)
    Dim lStyle As Long
    ShowWindow hwnd, SW_HIDE
    
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle
    
    App.TaskVisible = bShow
    
    ShowWindow hwnd, SW_NORMAL
End Sub

Public Property Get Caption() As String
    Caption = sFormCaption
End Property

Public Property Let Caption(ByVal New_TheCaption As String)
    sFormCaption = New_TheCaption
    PropertyChanged "Caption"
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get DisplayIcon() As Boolean
    DisplayIcon = bDisplayIcon
End Property

Public Property Let DisplayIcon(ByVal New_DisplayIcon As Boolean)
    bDisplayIcon = New_DisplayIcon
    PropertyChanged "DisplayIcon"
    Call UserControl_Resize
End Property

Public Property Get Font() As Font
    Set Font = picForm.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picForm.Font = New_Font
    bFontBold = picForm.FontBold
    bFontItalic = picForm.FontItalic
    dFontSize = picForm.FontSize
    bFontStrikeThru = picForm.FontStrikethru
    bFontUnderline = picForm.FontUnderline
    Call UserControl_Resize
    PropertyChanged "Font"
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picForm.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picForm.ForeColor() = New_ForeColor
    lFormCaptionColor = New_ForeColor
    PropertyChanged "ForeColor"
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get Icon() As Picture
    Set Icon = imgFormPic.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set imgFormPic.Picture = New_Icon
    Set myForm.Icon = New_Icon
    PropertyChanged "Icon"
    bDisplayIcon = True
    Call UserControl_Resize
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = lFormMaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Long)
    lFormMaxHeight = New_MaxHeight
    PropertyChanged "MaxHeight"
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = lFormMaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Long)
    lFormMaxWidth = New_MaxWidth
    PropertyChanged "MaxWidth"
End Property

Public Property Get MinHeight() As Long
    MinHeight = lFormMinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    lFormMinHeight = New_MinHeight
    PropertyChanged "MinHeight"
End Property

Public Property Get MinWidth() As Long
    MinWidth = lFormMinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    lFormMinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = bCloseButton
End Property

Public Property Let ShowCloseButton(ByVal New_ShowCloseButton As Boolean)
    bCloseButton = New_ShowCloseButton
    PropertyChanged "ShowCloseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowMinimiseButton() As Boolean
    ShowMinimiseButton = bMinimiseButton
End Property

Public Property Let ShowMinimiseButton(ByVal New_ShowMinimiseButton As Boolean)
    bMinimiseButton = New_ShowMinimiseButton
    PropertyChanged "ShowMinimiseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowMaximiseButton() As Boolean
    ShowMaximiseButton = bMaximiseButton
End Property

Public Property Let ShowMaximiseButton(ByVal New_ShowMaximiseButton As Boolean)
    bMaximiseButton = New_ShowMaximiseButton
    PropertyChanged "ShowMaximiseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowSystemTrayIcon() As Boolean
    ShowSystemTrayIcon = bSystemTray
End Property

Public Property Let ShowSystemTrayIcon(ByVal New_ShowSystemTrayIcon As Boolean)
    bSystemTray = New_ShowSystemTrayIcon
    PropertyChanged "ShowSystemTrayIcon"
    Call UserControl_Resize
End Property

Public Property Get Style() As xVistaStyles
    Style = xVisualStyles
End Property

Public Property Let Style(val As xVistaStyles)
    ' Determine which color scheme has been selected
    xVisualStyles = val

    ' Set the colour scheme
    Call SelectColorScheme

    picForm.ForeColor() = lFormCaptionColor
    lFormCaptionColor = lFormCaptionColor
    PropertyChanged "Style"

    ' Repaint the control
    Call UserControl_Paint

    ' Draw the Form Header and Buttons
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get Transparency() As Boolean
    Transparency = bTransparency
End Property

Public Property Let Transparency(ByVal New_Transparency As Boolean)
    bTransparency = New_Transparency
    PropertyChanged "Transparency"
    Call UserControl_Resize
End Property

Public Property Get TransparencyLevel() As String
    TransparencyLevel = iTransparency
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As String)
    iTransparency = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
    Call UserControl_Resize
End Property

Private Sub SelectColorScheme()
Select Case xVisualStyles
    Case VistaBlue
        lFormCaptionColor = &H0&
        lFormGradientBottom = &HEAD1B9
        lFormGradientTop = &HD0B498
        lFormInnerBorder = &HE7D3C1
        lFormMiddleBorder = &HE4CF28
        lFormOuterBorder = &H0&
        lButtonGradientBottom(1) = &HCCB198
        lButtonGradientBottom(2) = &HD1B79E
        lButtonGradientBottom(3) = &HD8BEA4
        lButtonGradientBottom(4) = &HDFC5AC
        lButtonGradientBottom(5) = &HDFC5AC
        lButtonGradientBottom(6) = &HE5CBB2
        lButtonGradientBottom(7) = &HE9D0B7
        lButtonGradientBottomClicked(1) = &H523B20
        lButtonGradientBottomClicked(2) = &H5B441F
        lButtonGradientBottomClicked(3) = &H736223
        lButtonGradientBottomClicked(4) = &H918727
        lButtonGradientBottomClicked(5) = &H918727
        lButtonGradientBottomClicked(6) = &HACA82B
        lButtonGradientBottomClicked(7) = &HC8C927
        lButtonGradientBottomHover(1) = &HA3732D
        lButtonGradientBottomHover(2) = &HAF7B2C
        lButtonGradientBottomHover(3) = &HBF892C
        lButtonGradientBottomHover(4) = &HD09A2C
        lButtonGradientBottomHover(5) = &HD09A2C
        lButtonGradientBottomHover(6) = &HDFA929
        lButtonGradientBottomHover(7) = &HEBC624
        lButtonGradientTop = &HE7D3C1
        lButtonGradientTopClicked = &H9C886E
        lButtonGradientTopHover = &HEFCB96
        lButtonInnerBorder = &HF2E7DE
        lButtonOuterBorder = &H886F5D
    Case VistaDark
        lFormCaptionColor = &HFFFFFF
        lFormGradientBottom = &H322624
        lFormGradientTop = &H9E9794
        lFormInnerBorder = &HE0E0E0
        lFormMiddleBorder = &HB9B8B4
        lFormOuterBorder = &H0&
        lButtonGradientBottom(1) = &H433E35
        lButtonGradientBottom(2) = &H464138
        lButtonGradientBottom(3) = &H413E36
        lButtonGradientBottom(4) = &H403D35
        lButtonGradientBottom(5) = &H403C37
        lButtonGradientBottom(6) = &H413D38
        lButtonGradientBottom(7) = &H423E39
        lButtonGradientBottomClicked(1) = &H2B2B2B
        lButtonGradientBottomClicked(2) = &H353535
        lButtonGradientBottomClicked(3) = &H414141
        lButtonGradientBottomClicked(4) = &H4F4F4F
        lButtonGradientBottomClicked(5) = &H4F4F4F
        lButtonGradientBottomClicked(6) = &H535353
        lButtonGradientBottomClicked(7) = &H666666
        lButtonGradientBottomHover(1) = &H7B695B
        lButtonGradientBottomHover(2) = &H816F60
        lButtonGradientBottomHover(3) = &H8C7B6C
        lButtonGradientBottomHover(4) = &H9A8979
        lButtonGradientBottomHover(5) = &H9A8979
        lButtonGradientBottomHover(6) = &HAA9A88
        lButtonGradientBottomHover(7) = &HBAAA97
        lButtonGradientTop = &HC0BDB8
        lButtonGradientTopClicked = &H4D4D4D
        lButtonGradientTopHover = &HD1CAC4
        lButtonInnerBorder = &HB9B8B4
        lButtonOuterBorder = &H221443
    Case VistaCustom
        lFormCaptionColor = &HFFFFFF
        lFormGradientBottom = xlFormGradientBottom
        lFormGradientTop = xlFormGradientTop
        lFormInnerBorder = xlFormInnerBorder
        lFormMiddleBorder = xlFormMiddleBorder
        lFormOuterBorder = xlFormOuterBorder
        lButtonGradientBottom(1) = xlButtonGradientBottom
        lButtonGradientBottom(2) = xlButtonGradientBottom
        lButtonGradientBottom(3) = xlButtonGradientBottom
        lButtonGradientBottom(4) = CreateGradientButom(10, 70, xlButtonGradientBottom)
        lButtonGradientBottom(5) = CreateGradientButom(10, 70, xlButtonGradientBottom)
        lButtonGradientBottom(6) = CreateGradientButom(20, 70, xlButtonGradientBottom)
        lButtonGradientBottom(7) = CreateGradientButom(25, 70, xlButtonGradientBottom)
        lButtonGradientBottomClicked(1) = xlButtonGradientBottomClicked
        lButtonGradientBottomClicked(2) = CreateGradientButom(10, 70, xlButtonGradientBottomClicked)
        lButtonGradientBottomClicked(3) = CreateGradientButom(15, 70, xlButtonGradientBottomClicked)
        lButtonGradientBottomClicked(4) = CreateGradientButom(15, 70, xlButtonGradientBottomClicked)
        lButtonGradientBottomClicked(5) = CreateGradientButom(20, 70, xlButtonGradientBottomClicked)
        lButtonGradientBottomClicked(6) = CreateGradientButom(25, 70, xlButtonGradientBottomClicked)
        lButtonGradientBottomClicked(7) = CreateGradientButom(30, 70, xlButtonGradientBottomClicked)
        lButtonGradientBottomHover(1) = xlButtonGradientBottomHover
        lButtonGradientBottomHover(2) = CreateGradientButom(10, 70, xlButtonGradientBottomHover)
        lButtonGradientBottomHover(3) = CreateGradientButom(15, 70, xlButtonGradientBottomHover)
        lButtonGradientBottomHover(4) = CreateGradientButom(15, 70, xlButtonGradientBottomHover)
        lButtonGradientBottomHover(5) = CreateGradientButom(20, 70, xlButtonGradientBottomHover)
        lButtonGradientBottomHover(6) = CreateGradientButom(25, 70, xlButtonGradientBottomHover)
        lButtonGradientBottomHover(7) = CreateGradientButom(30, 70, xlButtonGradientBottomHover)
        lButtonGradientTop = lFormInnerBorder
        lButtonGradientTopClicked = lFormInnerBorder
        lButtonGradientTopHover = lFormInnerBorder
        lButtonInnerBorder = CreateGradientButom(0, 70, xlButtonGradientBottomHover)
        lButtonOuterBorder = &H221443
End Select

lCloseButtonGradientBottom(1) = &H2C43B8
lCloseButtonGradientBottom(2) = &H3249BA
lCloseButtonGradientBottom(3) = &H3F54BF
lCloseButtonGradientBottom(4) = &H4F62C5
lCloseButtonGradientBottom(5) = &H4F62C5
lCloseButtonGradientBottom(6) = &H6373CD
lCloseButtonGradientBottom(7) = &H7685D5
lCloseButtonGradientBottomClicked(1) = &H1883&
lCloseButtonGradientBottomClicked(2) = &H1987&
lCloseButtonGradientBottomClicked(3) = &H12B85
lCloseButtonGradientBottomClicked(4) = &H124391
lCloseButtonGradientBottomClicked(5) = &H124391
lCloseButtonGradientBottomClicked(6) = &H2C68A8
lCloseButtonGradientBottomClicked(7) = &H4A93C1
lCloseButtonGradientBottomHover(1) = &H223D2
lCloseButtonGradientBottomHover(2) = &H223D2
lCloseButtonGradientBottomHover(3) = &HD33D5
lCloseButtonGradientBottomHover(4) = &H2151DA
lCloseButtonGradientBottomHover(5) = &H2151DA
lCloseButtonGradientBottomHover(6) = &H3974E0
lCloseButtonGradientBottomHover(7) = &H56A0E8
lCloseButtonGradientTop = &H929FE4
lCloseButtonGradientTopClicked = &H768FBF
lCloseButtonGradientTopHover = &HADB9FC

'lCloseButtonInnerBorder = lCloseButtonGradientBottomHover(1)
'lCloseButtonOuterBorder = &H221443
lCloseButtonInnerBorder = &HCCD3F4
lCloseButtonOuterBorder = &H221443

If Janela_Ativa = False Then
    lFormCaptionColor = CreateGradientButom(35, 70, lFormCaptionColor)
    lFormGradientBottom = CreateGradientButom(35, 70, lFormGradientBottom)
    lFormGradientTop = CreateGradientButom(35, 70, lFormGradientTop)
    lFormInnerBorder = CreateGradientButom(35, 70, lFormInnerBorder)
    lFormMiddleBorder = CreateGradientButom(35, 70, lFormMiddleBorder)
    lFormOuterBorder = CreateGradientButom(35, 70, lFormOuterBorder)
    lButtonGradientBottom(1) = CreateGradientButom(35, 70, lButtonGradientBottom(1))
    lButtonGradientBottom(2) = CreateGradientButom(35, 70, lButtonGradientBottom(2))
    lButtonGradientBottom(3) = CreateGradientButom(35, 70, lButtonGradientBottom(3))
    lButtonGradientBottom(4) = CreateGradientButom(35, 70, lButtonGradientBottom(4))
    lButtonGradientBottom(5) = CreateGradientButom(35, 70, lButtonGradientBottom(5))
    lButtonGradientBottom(6) = CreateGradientButom(35, 70, lButtonGradientBottom(6))
    lButtonGradientBottom(7) = CreateGradientButom(35, 70, lButtonGradientBottom(7))
    lButtonGradientBottomClicked(1) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(1))
    lButtonGradientBottomClicked(2) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(2))
    lButtonGradientBottomClicked(3) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(3))
    lButtonGradientBottomClicked(4) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(4))
    lButtonGradientBottomClicked(5) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(5))
    lButtonGradientBottomClicked(6) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(6))
    lButtonGradientBottomClicked(7) = CreateGradientButom(35, 70, lButtonGradientBottomClicked(7))
    lButtonGradientBottomHover(1) = CreateGradientButom(35, 70, lButtonGradientBottomHover(1))
    lButtonGradientBottomHover(2) = CreateGradientButom(35, 70, lButtonGradientBottomHover(2))
    lButtonGradientBottomHover(3) = CreateGradientButom(35, 70, lButtonGradientBottomHover(3))
    lButtonGradientBottomHover(4) = CreateGradientButom(35, 70, lButtonGradientBottomHover(4))
    lButtonGradientBottomHover(5) = CreateGradientButom(35, 70, lButtonGradientBottomHover(5))
    lButtonGradientBottomHover(6) = CreateGradientButom(35, 70, lButtonGradientBottomHover(6))
    lButtonGradientBottomHover(7) = CreateGradientButom(35, 70, lButtonGradientBottomHover(7))
    lButtonGradientTop = CreateGradientButom(35, 70, lButtonGradientTop)
    lButtonGradientTopClicked = CreateGradientButom(35, 70, lButtonGradientTopClicked)
    lButtonGradientTopHover = CreateGradientButom(35, 70, lButtonGradientTopHover)
    lButtonInnerBorder = CreateGradientButom(35, 70, lButtonInnerBorder)
    lButtonOuterBorder = CreateGradientButom(35, 70, lButtonOuterBorder)

    lCloseButtonGradientBottom(1) = lButtonGradientBottom(1)
    lCloseButtonGradientBottom(2) = lButtonGradientBottom(2)
    lCloseButtonGradientBottom(3) = lButtonGradientBottom(3)
    lCloseButtonGradientBottom(4) = lButtonGradientBottom(4)
    lCloseButtonGradientBottom(5) = lButtonGradientBottom(5)
    lCloseButtonGradientBottom(6) = lButtonGradientBottom(6)
    lCloseButtonGradientBottom(7) = lButtonGradientBottom(7)
    'lCloseButtonGradientBottom(1) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(1))
    'lCloseButtonGradientBottom(2) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(2))
    'lCloseButtonGradientBottom(3) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(3))
    'lCloseButtonGradientBottom(4) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(4))
    'lCloseButtonGradientBottom(5) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(5))
    'lCloseButtonGradientBottom(6) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(6))
    'lCloseButtonGradientBottom(7) = CreateGradientButom(35, 70, lCloseButtonGradientBottom(7))
    lCloseButtonGradientBottomClicked(1) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(1))
    lCloseButtonGradientBottomClicked(2) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(2))
    lCloseButtonGradientBottomClicked(3) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(3))
    lCloseButtonGradientBottomClicked(4) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(4))
    lCloseButtonGradientBottomClicked(5) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(5))
    lCloseButtonGradientBottomClicked(6) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(6))
    lCloseButtonGradientBottomClicked(7) = CreateGradientButom(35, 70, lCloseButtonGradientBottomClicked(7))
    lCloseButtonGradientBottomHover(1) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(1))
    lCloseButtonGradientBottomHover(2) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(2))
    lCloseButtonGradientBottomHover(3) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(3))
    lCloseButtonGradientBottomHover(4) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(4))
    lCloseButtonGradientBottomHover(5) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(5))
    lCloseButtonGradientBottomHover(6) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(6))
    lCloseButtonGradientBottomHover(7) = CreateGradientButom(35, 70, lCloseButtonGradientBottomHover(7))
    lCloseButtonGradientTop = lButtonGradientTop
    'lCloseButtonGradientTop = CreateGradientButom(35, 70, lCloseButtonGradientTop)
    lCloseButtonGradientTopClicked = CreateGradientButom(35, 70, lCloseButtonGradientTopClicked)
    lCloseButtonGradientTopHover = CreateGradientButom(35, 70, lCloseButtonGradientTopHover)
    
    lCloseButtonInnerBorder = lButtonInnerBorder
    lCloseButtonOuterBorder = lButtonOuterBorder
    'lCloseButtonInnerBorder = CreateGradientButom(35, 70, lCloseButtonInnerBorder)
    'lCloseButtonOuterBorder = CreateGradientButom(35, 70, lCloseButtonOuterBorder)
    
End If
End Sub

Private Sub UserControlsCreate()
    If iNumControls = 0 Then
        ' Create the controls only once
        iNumControls = 1
        
        ' Add the Form Header picturebox
        Set picForm = UserControl.Controls.Add("VB.PictureBox", "picForm")
        picForm.AutoRedraw = True
        picForm.BorderStyle = 0
        picForm.Visible = True
        
        ' Add the Form Timer
        Set TmrMouseMove = UserControl.Controls.Add("VB.Timer", "TmrMouseMove")
        TmrMouseMove.Enabled = False
        TmrMouseMove.Interval = 10
               
        ' Add the Form Header Image
        Set imgFormPic = Controls.Add("VB.Image", "imgFormPic", picForm)
        Set imgFormPic.Picture = Nothing
    End If
End Sub

Private Sub moveForm_Load()
    If bSystemTray = True Then
        Call AddIconToTray
    End If
    Dim Minimizar As Boolean
    If lFormMaxHeight > 0 Or lFormMaxWidth > 0 Then
        If lFormMaxHeight > 0 Then myForm.Height = lFormMaxHeight
        If lFormMaxWidth > 0 Then myForm.Width = lFormMaxWidth
    Else
        If myForm.Tag = "vbMaximized" Then
            myForm.Tag = "vbMaximized"
            FORMRECT.Top = myForm.Parent.Top
            FORMRECT.Left = myForm.Parent.Left
            FORMRECT.Width = myForm.Parent.Width
            FORMRECT.Height = myForm.Parent.Height
            myForm.Parent.Top = 30
            myForm.Parent.Left = 0
            myForm.Parent.Width = Screen.Width
            myForm.Parent.Height = Screen.Height - GetTaskbarHeight - 36
        Else
            myForm.Tag = "vbNormal"
            UserControl.Parent.Top = FORMRECT.Top
            UserControl.Parent.Left = FORMRECT.Left
            UserControl.Parent.Width = FORMRECT.Width
            UserControl.Parent.Height = FORMRECT.Height
        End If
    End If

    Call UserControl_Paint
End Sub

Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim lngReturnValue As Long
    
    If Button = vbLeftButton Then
        bMouseClicked = True
        If bMinimiseButtonHover = True Then
            bCloseButtonHover = False
            bCloseButtonClicked = False
            bMaximiseButtonClicked = False
            bMaximiseButtonHover = False
            bMinimiseButtonClicked = True
            Call DrawMinimiseButton
            picForm.Refresh
        ElseIf bMaximiseButtonHover = True Then
            bCloseButtonHover = False
            bCloseButtonClicked = False
            bMaximiseButtonClicked = True
            bMinimiseButtonClicked = False
            bMinimiseButtonHover = False
            Call DrawMaximiseButton
            picForm.Refresh
        ElseIf bCloseButtonHover = True Then
            bCloseButtonClicked = True
            bMaximiseButtonClicked = False
            bMaximiseButtonHover = False
            bMinimiseButtonClicked = False
            bMinimiseButtonHover = False
            Call DrawCloseButton
            picForm.Refresh
        End If
        
        If X < (UserControl.Width - 1500) And (Y > 30 And Y <= 375) Then
            Call ReleaseCapture
            myForm.MousePointer = vbSizeAll
            lngReturnValue = SendMessage(moveForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
            myForm.MousePointer = vbNormal
        End If
        
        Dim RLeft As Single
        Dim RTop As Single
        Dim RRight As Boolean
        Dim RBottom As Boolean
        Dim StartX As Single
        Dim StartY As Single
        Dim MX As Single
        Dim MY As Single

        ' Read the mouse pointer screen position on the beginning
        GetMouseXY StartX, StartY
        ' We don't use the X,Y arguments which are incorect when MouseDown event called
        ' from other screen objects (like Label1 and Image1 hereunder)
        X = StartX - myForm.Left
        Y = StartY - myForm.Top

        ' Flags indicating "from where" the form is being resized
        RLeft = IIf(X < BorderWidth, myForm.Width, 0)
        RTop = IIf(Y < BorderHeight, myForm.Height, 0)
        RRight = (X > myForm.Width - BorderWidth)
        RBottom = (Y > myForm.Height - BorderHeight)

        ' Place the mouse pointer on the form border for more accuracy
        If RLeft Then SetMouseXY myForm.Left, myForm.Top + Y
        If RTop Then SetMouseXY myForm.Left + X, myForm.Top
        If RRight Then SetMouseXY myForm.Left + myForm.Width, myForm.Top + Y
        If RBottom Then SetMouseXY myForm.Left + X, myForm.Top + myForm.Height

        ' Save the mouse pointer screen position on the beginning in variables
        GetMouseXY StartX, StartY
        ' While left mouse button is pressed
        While GetAsyncKeyState(vbLeftButton) < 0
            ' Read the actual mouse pointer screen position
            GetMouseXY MX, MY

            If RRight Or RLeft Or RBottom Or RTop Then      ' If the form is resized (not moved)
                If lFormMaxHeight > 0 And myForm.Height - 5 > lFormMaxHeight Then
                    myForm.Height = lFormMaxHeight
                    Exit Sub
                End If
                If lFormMaxWidth > 0 And myForm.Width - 5 > lFormMaxWidth Then
                    myForm.Width = lFormMaxWidth
                    Exit Sub
                End If
                If lFormMinHeight > 0 And myForm.Height + 5 < lFormMinHeight Then
                    myForm.Height = lFormMinHeight
                    Exit Sub
                End If
                If lFormMinWidth > 0 And myForm.Width + 5 < lFormMinWidth Then
                    myForm.Width = lFormMinWidth
                    Exit Sub
                End If
                If RLeft And RLeft + StartX - MX > BorderWidth * 2 Then myForm.Move MX, myForm.Top, RLeft + StartX - MX
                If RTop And RTop + StartY - MY > BorderHeight * 2 Then myForm.Move myForm.Left, MY, myForm.Width, RTop + StartY - MY
                If RRight And MX - myForm.Left > BorderWidth * 2 Then myForm.Width = MX - myForm.Left
                If RBottom And MY - myForm.Top > BorderHeight * 2 Then myForm.Height = MY - myForm.Top
''                Else                                            ' If the form is moved (not resized)
''                    MousePointer = vbSizeAll                    ' Sets the mouse cursor showing move
''                    myForm.Move MX - X, MY - Y                  ' Actually moves the form on screen
            End If
            DoEvents                                        ' To allow Windows painting events
        Wend
        
        If bUnloadForm = False Then
            Call DrawMinimiseButton
            Call DrawMaximiseButton
            Call DrawCloseButton
            myForm.Refresh
        End If
    End If
End Sub

Private Sub moveForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Reset the mouse cursor when left mouse button is not pressed
    If GetAsyncKeyState(vbLeftButton) >= 0 Then myForm.MousePointer = 0
    
    ' Set the correct mouse cursor according to its position on the form
    If (X + BorderWidth) > myForm.Width - BorderWidth Or X < BorderWidth Then myForm.MousePointer = vbSizeWE
    If bRightClick = False And (Y > myForm.Height - BorderHeight Or Y < 30) Then myForm.MousePointer = vbSizeNS
    If Y > 30 And Y <= (UserControl.Height + 15) And (X > 30 And X < UserControl.Width - 30) Then myForm.MousePointer = vbNormal
    If (X + BorderWidth > myForm.Width - BorderWidth And Y > myForm.Height - BorderHeight) Or (X < BorderWidth And Y < BorderHeight) Then myForm.MousePointer = vbSizeNWSE
    If (X + BorderWidth > myForm.Width - BorderWidth And Y < BorderHeight) Or (X < BorderWidth And Y > myForm.Height - BorderHeight) Then myForm.MousePointer = vbSizeNESW
End Sub

Private Sub moveForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMinimiseButtonHover = True And bMinimiseButton = True And bMinimiseButtonClicked = True Then
        If bEnableMinimiseButton Then
            myForm.Tag = "vbMinimized"
            myForm.WindowState = vbMinimized
        End If
    ElseIf bMaximiseButtonHover = True And bMaximiseButton = True And bMaximiseButtonClicked = True Then
        If myForm.Tag <> "vbMaximized" Then
            If lFormMaxHeight > 0 Or lFormMaxWidth > 0 Then
                If lFormMaxHeight > 0 Then myForm.Height = lFormMaxHeight
                If lFormMaxWidth > 0 Then myForm.Width = lFormMaxWidth
            Else
                If bEnableMaximiseButton Then
                    myForm.Tag = "vbMaximized"
                    FORMRECT.Top = UserControl.Parent.Top
                    FORMRECT.Left = UserControl.Parent.Left
                    FORMRECT.Width = UserControl.Parent.Width
                    FORMRECT.Height = UserControl.Parent.Height
                    UserControl.Parent.Top = 30
                    UserControl.Parent.Left = 0
                    UserControl.Parent.Width = Screen.Width
                    UserControl.Parent.Height = Screen.Height - GetTaskbarHeight - 36
                End If
            End If
        Else
            myForm.Tag = "vbNormal"
            UserControl.Parent.Top = FORMRECT.Top
            UserControl.Parent.Left = FORMRECT.Left
            UserControl.Parent.Width = FORMRECT.Width
            UserControl.Parent.Height = FORMRECT.Height
        End If
    ElseIf bCloseButtonHover = True And bCloseButton = True And bCloseButtonClicked = True Then
        If bEnableCloseButton = True Then
            TmrMouseMove.Enabled = False
            bUnloadForm = True
            Unload moveForm
            Unload myForm
        End If
    Else
        bCloseButtonClicked = False
        bMaximiseButtonClicked = False
        bMinimiseButtonClicked = False
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
        picForm.Refresh
    End If
    bMouseClicked = False
    bMouseOnForm = False
End Sub

' Reads actual mouse pointer screen position and convert it to TWIP scale
Private Sub GetMouseXY(X As Single, Y As Single)
    Dim lpPoint As POINTAPI

    GetCursorPos lpPoint
    X = lpPoint.X * TwipX
    Y = lpPoint.Y * TwipY
End Sub

' Places mouse pointer on given screen position given in TWIP scale
Private Sub SetMouseXY(ByVal X As Single, ByVal Y As Single)
    SetCursorPos X / TwipX, Y / TwipY
End Sub

Private Sub moveForm_Resize()
    Call UserControl_Resize
    myForm.Refresh
End Sub

Private Sub moveForm_Terminate()
    Call DeleteIconFromTray
End Sub

Private Sub moveForm_Unload(Cancel As Integer)
    Call DeleteIconFromTray
  
    ' Destroy the menu before exiting
    If lSysTrayMenu Then
      Call DestroyMenu(lSysTrayMenu)
      bUnloadForm = True
    End If
End Sub

Private Sub picForm_DblClick()
    If myForm.Tag <> "vbMaximized" And bMaximiseButton = True Then
        If lFormMaxHeight > 0 Or lFormMaxWidth > 0 Then
            If lFormMaxHeight > 0 Then myForm.Height = lFormMaxHeight
            If lFormMaxWidth > 0 Then myForm.Width = lFormMaxWidth
        Else
            If bEnableMaximiseButton Then
                myForm.Tag = "vbMaximized"
                FORMRECT.Top = UserControl.Parent.Top
                FORMRECT.Left = UserControl.Parent.Left
                FORMRECT.Width = UserControl.Parent.Width
                FORMRECT.Height = UserControl.Parent.Height
                UserControl.Parent.Width = Screen.Width
                UserControl.Parent.Height = Screen.Height - GetTaskbarHeight - 36
                UserControl.Parent.Top = 30
                UserControl.Parent.Left = 0
            End If
        End If
    Else
        myForm.Tag = "vbNormal"
        UserControl.Parent.Top = FORMRECT.Top
        UserControl.Parent.Left = FORMRECT.Left
        UserControl.Parent.Width = FORMRECT.Width
        UserControl.Parent.Height = FORMRECT.Height
    End If
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Style_Type = Vista_Aero Then
    ' Set the mouse to default

            If Y > 30 And (X > 30 Or X < UserControl.Width - 30) Then
                myForm.MousePointer = vbNormal
            End If
            
            ' Determine which button has hover focus
            If ((X >= (UserControl.Width - 1470)) And (X <= (UserControl.Width - 1065))) And (Y >= 30 And Y <= 240) Then
                
                bCloseButtonHover = False
                bMaximiseButtonHover = False
                bMinimiseButtonHover = True
                If Button = 1 Then
                    If bMouseClicked = False Then
                        bMinimiseButtonClicked = True
                    End If
                End If
            ElseIf ((X >= (UserControl.Width - 1035)) And (X <= (UserControl.Width - 630))) And ((Y >= 30) And (Y <= 240)) Then
                
                bCloseButtonHover = False
                bMaximiseButtonHover = True
                bMinimiseButtonHover = False
                If Button = 1 Then
                    If bMouseClicked = False Then
                        bMaximiseButtonClicked = True
                    End If
                End If
            ElseIf ((X >= (UserControl.Width - 1035)) And (X <= (UserControl.Width - 120))) And ((Y >= 30) And (Y <= 240)) Then
                'Debug.Print
                'Debug.Print UserControl.Width - 8160
                'Debug.Print UserControl.Width - 9075
                
                bCloseButtonHover = True
                bMaximiseButtonHover = False
                bMinimiseButtonHover = False
                If Button = 1 Then
                    If bMouseClicked = False Then
                        bCloseButtonClicked = True
                    End If
                End If
            Else
                bCloseButtonClicked = False
                bCloseButtonHover = False
                bMaximiseButtonClicked = False
                bMaximiseButtonHover = False
                bMinimiseButtonClicked = False
                bMinimiseButtonHover = False
            End If
            
            ' Enable the timer
            TmrMouseMove.Enabled = True
            Call moveForm_MouseMove(Button, Shift, X, Y)
    Else
    ' Set the mouse to default
            If Y > 30 And (X > 30 Or X < UserControl.Width - 30) Then myForm.MousePointer = vbNormal
            
            ' Determine which button has hover focus
            If (X >= (UserControl.Width - 1470) And X <= (UserControl.Width - 1050)) And (Y >= 75 And Y <= 300) Then
                bCloseButtonHover = False
                bMaximiseButtonHover = False
                bMinimiseButtonHover = True
                If Button = 1 Then
                    If bMouseClicked = False Then
                        bMinimiseButtonClicked = True
                    End If
                End If
            ElseIf X >= (UserControl.Width - 1005) And X <= (UserControl.Width - 575) And (Y >= 75 And Y <= 300) Then
                bCloseButtonHover = False
                bMaximiseButtonHover = True
                bMinimiseButtonHover = False
                If Button = 1 Then
                    If bMouseClicked = False Then
                        bMaximiseButtonClicked = True
                    End If
                End If
            ElseIf X >= (UserControl.Width - 520) And X <= (UserControl.Width - 105) And (Y >= 75 And Y <= 300) Then
                bCloseButtonHover = True
                bMaximiseButtonHover = False
                bMinimiseButtonHover = False
                If Button = 1 Then
                    If bMouseClicked = False Then
                        bCloseButtonClicked = True
                    End If
                End If
            Else
                bCloseButtonClicked = False
                bCloseButtonHover = False
                bMaximiseButtonClicked = False
                bMaximiseButtonHover = False
                bMinimiseButtonClicked = False
                bMinimiseButtonHover = False
            End If
            
            ' Enable the timer
            TmrMouseMove.Enabled = True
            
            Call moveForm_MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub DrawCloseButton()
'On Error GoTo Errhandler
Dim Vi As Integer
Dim VWidth As Long
Dim VTop As Long

VWidth = UserControl.Width
VTop = UserControl.Extender.Top

    If Style_Type = Vista_Aero Then
        If bCloseButton = True Then
            If bEnableCloseButton Then
                ' Draw the Buttons Outside Border
                picForm.Line (VWidth - 510 - 90, VTop + 75 - 60)-(VWidth - 105, VTop + 75 - 60), lCloseButtonOuterBorder
                picForm.Line (VWidth - 510 - 90, VTop + 315 - 60)-(VWidth - 105 - 30, VTop + 315 - 60), lCloseButtonOuterBorder
                picForm.Line (VWidth - 525 - 90, VTop + 90 - 60)-(VWidth - 525 - 90, VTop + 315 - 60), lCloseButtonOuterBorder
                picForm.Line (VWidth - 105, VTop + 90 - 60)-(VWidth - 105, VTop + 315 - 60 - 30), lCloseButtonOuterBorder
                picForm.Line (VWidth - 105, VTop + 215)-(VWidth - 145, VTop + 260), lCloseButtonOuterBorder
                
                picForm.Line (VWidth - 495 - 90, VTop + 90 - 60)-(VWidth - 120, VTop + 90 - 60), lCloseButtonInnerBorder
                picForm.Line (VWidth - 510 - 90, VTop + 300 - 60)-(VWidth - 105 - 45, VTop + 300 - 60), lCloseButtonInnerBorder
                picForm.Line (VWidth - 510 - 90, VTop + 90 - 60)-(VWidth - 510 - 90, VTop + 300 - 60), lCloseButtonInnerBorder
                picForm.Line (VWidth - 120, VTop + 90 - 60)-(VWidth - 120, VTop + 300 - 60 - 30), lCloseButtonInnerBorder
                picForm.Line (VWidth - 120, VTop + 215)-(VWidth - 160, VTop + 260), lCloseButtonInnerBorder
                
                If bCloseButtonHover = True And bCloseButtonClicked = False Then
                    picForm.Line (VWidth - 495 - 90, VTop + 90 - 60)-(VWidth - 120, VTop + 90 - 60), 13423604
                    picForm.Line (VWidth - 510 - 90, VTop + 300 - 60)-(VWidth - 105 - 45, VTop + 300 - 60), 13423604
                    picForm.Line (VWidth - 510 - 90, VTop + 90 - 60)-(VWidth - 510 - 90, VTop + 300 - 60), 13423604
                    picForm.Line (VWidth - 120, VTop + 90 - 60)-(VWidth - 120, VTop + 300 - 60 - 30), 13423604
                    picForm.Line (VWidth - 120, VTop + 215)-(VWidth - 160, VTop + 260), 13423604
                End If
                
                If bCloseButtonHover = True And bCloseButtonClicked = True Then
                    picForm.Line (VWidth - 495 - 90, VTop + 90 - 60)-(VWidth - 120, VTop + 90 - 60), 13423604
                    picForm.Line (VWidth - 510 - 90, VTop + 300 - 60)-(VWidth - 105 - 45, VTop + 300 - 60), 13423604
                    picForm.Line (VWidth - 510 - 90, VTop + 90 - 60)-(VWidth - 510 - 90, VTop + 300 - 60), 13423604
                    picForm.Line (VWidth - 120, VTop + 90 - 60)-(VWidth - 120, VTop + 300 - 60 - 30), 13423604
                    picForm.Line (VWidth - 120, VTop + 215)-(VWidth - 160, VTop + 260), 13423604
                End If
                
                ' Close Button Top Gradient Base Colour
                iVertical = 105
                For I = 1 To 6
                    If bCloseButtonHover = False And bCloseButtonClicked = False Then
                        picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120, VTop + iVertical - 60), lCloseButtonGradientTop
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                        picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120, VTop + iVertical - 60), lCloseButtonGradientTopHover
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                        picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120, VTop + iVertical - 60), lCloseButtonGradientTopClicked
                    End If
                    iVertical = iVertical + 15
                Next
                
                ' Button Bottom Gradient Base Colour
                Vi = 0
                For I = 1 To 7
                    If I >= 7 Then
                        Vi = Vi + 15
                    End If
                    If bCloseButtonHover = False And bCloseButtonClicked = False Then
                        picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120 - Vi, VTop + iVertical - 60), lCloseButtonGradientBottom(I)
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                        picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120 - Vi, VTop + iVertical - 60), lCloseButtonGradientBottomHover(I)
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                        picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120 - Vi, VTop + iVertical - 60), lCloseButtonGradientBottomClicked(I)
                    End If
                    iVertical = iVertical + 15
                Next
                
                ' Draw the Close Button Display
                ' Outside Borders
                picForm.Line (VWidth - 375 - 45, VTop + 135 - 60)-(VWidth - 330 - 45, VTop + 135 - 60), &H665653
                picForm.Line (VWidth - 285 - 45, VTop + 135 - 60)-(VWidth - 240 - 45, VTop + 135 - 60), &H665653
                
                picForm.Line (VWidth - 390 - 45, VTop + 150 - 60)-(VWidth - 375 - 45, VTop + 150 - 60), &H665653
                picForm.Line (VWidth - 330 - 45, VTop + 150 - 60)-(VWidth - 315 - 45, VTop + 150 - 60), &H665653
                picForm.Line (VWidth - 300 - 45, VTop + 150 - 60)-(VWidth - 285 - 45, VTop + 150 - 60), &H665653
                picForm.Line (VWidth - 240 - 45, VTop + 150 - 60)-(VWidth - 225 - 45, VTop + 150 - 60), &H665653
                
                picForm.Line (VWidth - 375 - 45, VTop + 165 - 60)-(VWidth - 360 - 45, VTop + 165 - 60), &H665653
                picForm.Line (VWidth - 315 - 45, VTop + 165 - 60)-(VWidth - 300 - 45, VTop + 165 - 60), &H665653
                picForm.Line (VWidth - 255 - 45, VTop + 165 - 60)-(VWidth - 240 - 45, VTop + 165 - 60), &H665653
               
                picForm.Line (VWidth - 360 - 45, VTop + 180 - 60)-(VWidth - 345 - 45, VTop + 180 - 60), &H665653
                picForm.Line (VWidth - 270 - 45, VTop + 180 - 60)-(VWidth - 255 - 45, VTop + 180 - 60), &H665653
        
                picForm.Line (VWidth - 345 - 45, VTop + 195 - 60)-(VWidth - 330 - 45, VTop + 195 - 60), &H665653
                picForm.Line (VWidth - 285 - 45, VTop + 195 - 60)-(VWidth - 270 - 45, VTop + 195 - 60), &H665653
        
                picForm.Line (VWidth - 360 - 45, VTop + 210 - 60)-(VWidth - 345 - 45, VTop + 210 - 60), &H665653
                picForm.Line (VWidth - 270 - 45, VTop + 210 - 60)-(VWidth - 255 - 45, VTop + 210 - 60), &H665653
        
                picForm.Line (VWidth - 375 - 45, VTop + 225 - 60)-(VWidth - 360 - 45, VTop + 225 - 60), &H665653
                picForm.Line (VWidth - 315 - 45, VTop + 225 - 60)-(VWidth - 300 - 45, VTop + 225 - 60), &H665653
                picForm.Line (VWidth - 255 - 45, VTop + 225 - 60)-(VWidth - 240 - 45, VTop + 225 - 60), &H665653
        
                picForm.Line (VWidth - 390 - 45, VTop + 240 - 60)-(VWidth - 375 - 45, VTop + 240 - 60), &H665653
                picForm.Line (VWidth - 330 - 45, VTop + 240 - 60)-(VWidth - 315 - 45, VTop + 240 - 60), &H665653
                picForm.Line (VWidth - 300 - 45, VTop + 240 - 60)-(VWidth - 285 - 45, VTop + 240 - 60), &H665653
                picForm.Line (VWidth - 240 - 45, VTop + 240 - 60)-(VWidth - 225 - 45, VTop + 240 - 60), &H665653
        
                picForm.Line (VWidth - 375 - 45, VTop + 255 - 60)-(VWidth - 330 - 45, VTop + 255 - 60), &H665653
                picForm.Line (VWidth - 285 - 45, VTop + 255 - 60)-(VWidth - 240 - 45, VTop + 255 - 60), &H665653
        
                ' Inside Button Colours
                picForm.Line (VWidth - 375 - 45, VTop + 150 - 60)-(VWidth - 330 - 45, VTop + 150 - 60), &HFFFFFF
                picForm.Line (VWidth - 285 - 45, VTop + 150 - 60)-(VWidth - 240 - 45, VTop + 150 - 60), &HFFFFFF
                
                picForm.Line (VWidth - 360 - 45, VTop + 165 - 60)-(VWidth - 315 - 45, VTop + 165 - 60), &HFFFFFF
                picForm.Line (VWidth - 300 - 45, VTop + 165 - 60)-(VWidth - 255 - 45, VTop + 165 - 60), &HFFFFFF
                
                picForm.Line (VWidth - 345 - 45, VTop + 180 - 60)-(VWidth - 270 - 45, VTop + 180 - 60), &HFFFFFF
                
                picForm.Line (VWidth - 330 - 45, VTop + 195 - 60)-(VWidth - 285 - 45, VTop + 195 - 60), &HE9E9E9
                
                picForm.Line (VWidth - 345 - 45, VTop + 210 - 60)-(VWidth - 270 - 45, VTop + 210 - 60), &HE2E2E2
                
                picForm.Line (VWidth - 360 - 45, VTop + 225 - 60)-(VWidth - 315 - 45, VTop + 225 - 60), &HDCDCDC
                picForm.Line (VWidth - 300 - 45, VTop + 225 - 60)-(VWidth - 255 - 45, VTop + 225 - 60), &HDCDCDC
                
                picForm.Line (VWidth - 375 - 45, VTop + 240 - 60)-(VWidth - 330 - 45, VTop + 240 - 60), &HD7D7D7
                picForm.Line (VWidth - 285 - 45, VTop + 240 - 60)-(VWidth - 240 - 45, VTop + 240 - 60), &HD7D7D7
            Else
                ' Draw the Buttons Outside Border
                picForm.Line (VWidth - 510 - 90, VTop + 75 - 60)-(VWidth - 105, VTop + 75 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (VWidth - 510 - 90, VTop + 315 - 60)-(VWidth - 105 - 30, VTop + 315 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
                'picForm.Line (VWidth - 525 - 90, VTop + 90 - 60)-(VWidth - 525 - 90, VTop + 315 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (VWidth - 105, VTop + 90 - 60)-(VWidth - 105, VTop + 315 - 60 - 30), CreateGradientButom(35, 70, lButtonOuterBorder)
                
                picForm.Line (VWidth - 105, VTop + 215)-(VWidth - 145, VTop + 260), CreateGradientButom(35, 70, lButtonOuterBorder)
                
                picForm.Line (VWidth - 495 - 90, VTop + 90 - 60)-(VWidth - 120, VTop + 90 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (VWidth - 510 - 90, VTop + 300 - 60)-(VWidth - 105 - 45, VTop + 300 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (VWidth - 510 - 90, VTop + 90 - 60)-(VWidth - 510 - 90, VTop + 300 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (VWidth - 120, VTop + 90 - 60)-(VWidth - 120, VTop + 300 - 60 - 30), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (VWidth - 120, VTop + 215)-(VWidth - 160, VTop + 260), CreateGradientButom(35, 70, lButtonInnerBorder)
                
                ' Close Button Top Gradient Base Colour
                iVertical = 105
                For I = 1 To 6
                    picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120, VTop + iVertical - 60), CreateGradientButom(35, 70, lButtonGradientTop)
                    iVertical = iVertical + 15
                Next
                
                ' Button Bottom Gradient Base Colour
                Vi = 0
                For I = 1 To 7
                    If I >= 7 Then
                        Vi = Vi + 15
                    End If
                    picForm.Line (VWidth - 495 - 90, VTop + iVertical - 60)-(VWidth - 120 - Vi, VTop + iVertical - 60), CreateGradientButom(35, 70, lButtonGradientBottom(I))
                    iVertical = iVertical + 15
                Next
                
                ' Draw the Close Button Display
                ' Outside Borders
                picForm.Line (VWidth - 375 - 45, VTop + 135 - 60)-(VWidth - 330 - 45, VTop + 135 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 285 - 45, VTop + 135 - 60)-(VWidth - 240 - 45, VTop + 135 - 60), CreateGradientButom(35, 70, &H665653)
                
                picForm.Line (VWidth - 390 - 45, VTop + 150 - 60)-(VWidth - 375 - 45, VTop + 150 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 330 - 45, VTop + 150 - 60)-(VWidth - 315 - 45, VTop + 150 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 300 - 45, VTop + 150 - 60)-(VWidth - 285 - 45, VTop + 150 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 240 - 45, VTop + 150 - 60)-(VWidth - 225 - 45, VTop + 150 - 60), CreateGradientButom(35, 70, &H665653)
                
                picForm.Line (VWidth - 375 - 45, VTop + 165 - 60)-(VWidth - 360 - 45, VTop + 165 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 315 - 45, VTop + 165 - 60)-(VWidth - 300 - 45, VTop + 165 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 255 - 45, VTop + 165 - 60)-(VWidth - 240 - 45, VTop + 165 - 60), CreateGradientButom(35, 70, &H665653)
               
                picForm.Line (VWidth - 360 - 45, VTop + 180 - 60)-(VWidth - 345 - 45, VTop + 180 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 270 - 45, VTop + 180 - 60)-(VWidth - 255 - 45, VTop + 180 - 60), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 345 - 45, VTop + 195 - 60)-(VWidth - 330 - 45, VTop + 195 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 285 - 45, VTop + 195 - 60)-(VWidth - 270 - 45, VTop + 195 - 60), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 360 - 45, VTop + 210 - 60)-(VWidth - 345 - 45, VTop + 210 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 270 - 45, VTop + 210 - 60)-(VWidth - 255 - 45, VTop + 210 - 60), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 375 - 45, VTop + 225 - 60)-(VWidth - 360 - 45, VTop + 225 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 315 - 45, VTop + 225 - 60)-(VWidth - 300 - 45, VTop + 225 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 255 - 45, VTop + 225 - 60)-(VWidth - 240 - 45, VTop + 225 - 60), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 390 - 45, VTop + 240 - 60)-(VWidth - 375 - 45, VTop + 240 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 330 - 45, VTop + 240 - 60)-(VWidth - 315 - 45, VTop + 240 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 300 - 45, VTop + 240 - 60)-(VWidth - 285 - 45, VTop + 240 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 240 - 45, VTop + 240 - 60)-(VWidth - 225 - 45, VTop + 240 - 60), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 375 - 45, VTop + 255 - 60)-(VWidth - 330 - 45, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 285 - 45, VTop + 255 - 60)-(VWidth - 240 - 45, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653)
        
                ' Inside Button Colours
                picForm.Line (VWidth - 375 - 45, VTop + 150 - 60)-(VWidth - 330 - 45, VTop + 150 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                picForm.Line (VWidth - 285 - 45, VTop + 150 - 60)-(VWidth - 240 - 45, VTop + 150 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                
                picForm.Line (VWidth - 360 - 45, VTop + 165 - 60)-(VWidth - 315 - 45, VTop + 165 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                picForm.Line (VWidth - 300 - 45, VTop + 165 - 60)-(VWidth - 255 - 45, VTop + 165 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                
                picForm.Line (VWidth - 345 - 45, VTop + 180 - 60)-(VWidth - 270 - 45, VTop + 180 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                
                picForm.Line (VWidth - 330 - 45, VTop + 195 - 60)-(VWidth - 285 - 45, VTop + 195 - 60), CreateGradientButom(35, 70, &HE9E9E9)
                
                picForm.Line (VWidth - 345 - 45, VTop + 210 - 60)-(VWidth - 270 - 45, VTop + 210 - 60), CreateGradientButom(35, 70, &HE2E2E2)
                
                picForm.Line (VWidth - 360 - 45, VTop + 225 - 60)-(VWidth - 315 - 45, VTop + 225 - 60), CreateGradientButom(35, 70, &HDCDCDC)
                picForm.Line (VWidth - 300 - 45, VTop + 225 - 60)-(VWidth - 255 - 45, VTop + 225 - 60), CreateGradientButom(35, 70, &HDCDCDC)
                
                picForm.Line (VWidth - 375 - 45, VTop + 240 - 60)-(VWidth - 330 - 45, VTop + 240 - 60), CreateGradientButom(35, 70, &HD7D7D7)
                picForm.Line (VWidth - 285 - 45, VTop + 240 - 60)-(VWidth - 240 - 45, VTop + 240 - 60), CreateGradientButom(35, 70, &HD7D7D7)
            End If
        End If
    Else
        If bCloseButton = True Then
            If bEnableCloseButton = False Then
                ' Draw the Buttons Outside Border
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 75)-(UserControl.Width - 105, UserControl.Extender.Top + 75), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 315)-(UserControl.Width - 105, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 90)-(UserControl.Width - 525, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 90)-(UserControl.Width - 105, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
        
                ' Draw the Buttons Inside Border
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 90), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 300)-(UserControl.Width - 105, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 90)-(UserControl.Width - 510, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (UserControl.Width - 120, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
        
                ' Close Button Top Gradient Base Colour
                iVertical = 105
                For I = 1 To 6
                    picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), CreateGradientButom(35, 70, lButtonGradientTop)
                    iVertical = iVertical + 15
                Next
                
                ' Button Bottom Gradient Base Colour
                For I = 1 To 7
                    picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), CreateGradientButom(35, 70, lButtonGradientBottom(I))
                    iVertical = iVertical + 15
                Next
                
                ' Draw the Close Button Display
                ' Outside Borders
                picForm.Line (VWidth - 375, VTop + 135)-(VWidth - 330, VTop + 135), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 285, VTop + 135)-(VWidth - 240, VTop + 135), CreateGradientButom(35, 70, &H665653)
                
                picForm.Line (VWidth - 390, VTop + 150)-(VWidth - 375, VTop + 150), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 330, VTop + 150)-(VWidth - 315, VTop + 150), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 300, VTop + 150)-(VWidth - 285, VTop + 150), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 240, VTop + 150)-(VWidth - 225, VTop + 150), CreateGradientButom(35, 70, &H665653)
                
                picForm.Line (VWidth - 375, VTop + 165)-(VWidth - 360, VTop + 165), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 315, VTop + 165)-(VWidth - 300, VTop + 165), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 255, VTop + 165)-(VWidth - 240, VTop + 165), CreateGradientButom(35, 70, &H665653)
               
                picForm.Line (VWidth - 360, VTop + 180)-(VWidth - 345, VTop + 180), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 270, VTop + 180)-(VWidth - 255, VTop + 180), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 345, VTop + 195)-(VWidth - 330, VTop + 195), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 285, VTop + 195)-(VWidth - 270, VTop + 195), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 360, VTop + 210)-(VWidth - 345, VTop + 210), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 270, VTop + 210)-(VWidth - 255, VTop + 210), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 375, VTop + 225)-(VWidth - 360, VTop + 225), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 315, VTop + 225)-(VWidth - 300, VTop + 225), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 255, VTop + 225)-(VWidth - 240, VTop + 225), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 390, VTop + 240)-(VWidth - 375, VTop + 240), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 330, VTop + 240)-(VWidth - 315, VTop + 240), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 300, VTop + 240)-(VWidth - 285, VTop + 240), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 240, VTop + 240)-(VWidth - 225, VTop + 240), CreateGradientButom(35, 70, &H665653)
        
                picForm.Line (VWidth - 375, VTop + 255)-(VWidth - 330, VTop + 255), CreateGradientButom(35, 70, &H665653)
                picForm.Line (VWidth - 285, VTop + 255)-(VWidth - 240, VTop + 255), CreateGradientButom(35, 70, &H665653)
        
                ' Inside Button Colours
                picForm.Line (VWidth - 375, VTop + 150)-(VWidth - 330, VTop + 150), CreateGradientButom(35, 70, &HFFFFFF)
                picForm.Line (VWidth - 285, VTop + 150)-(VWidth - 240, VTop + 150), CreateGradientButom(35, 70, &HFFFFFF)
                
                picForm.Line (VWidth - 360, VTop + 165)-(VWidth - 315, VTop + 165), CreateGradientButom(35, 70, &HFFFFFF)
                picForm.Line (VWidth - 300, VTop + 165)-(VWidth - 255, VTop + 165), CreateGradientButom(35, 70, &HFFFFFF)
                
                picForm.Line (VWidth - 345, VTop + 180)-(VWidth - 270, VTop + 180), CreateGradientButom(35, 70, &HFFFFFF)
                
                picForm.Line (VWidth - 330, VTop + 195)-(VWidth - 285, VTop + 195), CreateGradientButom(35, 70, &HE9E9E9)
                
                picForm.Line (VWidth - 345, VTop + 210)-(VWidth - 270, VTop + 210), CreateGradientButom(35, 70, &HE2E2E2)
                
                picForm.Line (VWidth - 360, VTop + 225)-(VWidth - 315, VTop + 225), CreateGradientButom(35, 70, &HDCDCDC)
                picForm.Line (VWidth - 300, VTop + 225)-(VWidth - 255, VTop + 225), CreateGradientButom(35, 70, &HDCDCDC)
                
                picForm.Line (VWidth - 375, VTop + 240)-(VWidth - 330, VTop + 240), CreateGradientButom(35, 70, &HD7D7D7)
                picForm.Line (VWidth - 285, VTop + 240)-(VWidth - 240, VTop + 240), CreateGradientButom(35, 70, &HD7D7D7)
            Else
                ' Draw the Buttons Outside Border
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 75)-(UserControl.Width - 105, UserControl.Extender.Top + 75), lCloseButtonOuterBorder
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 315)-(UserControl.Width - 105, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
                picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 90)-(UserControl.Width - 525, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
                picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 90)-(UserControl.Width - 105, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
        
                ' Draw the Buttons Inside Border
                
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 90), lCloseButtonInnerBorder
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 300)-(UserControl.Width - 105, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
                picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 90)-(UserControl.Width - 510, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
                picForm.Line (UserControl.Width - 120, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
        
                If bCloseButtonHover = True And bCloseButtonClicked = False Then
                    picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 90), 13423604
                    picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 300)-(UserControl.Width - 105, UserControl.Extender.Top + 300), 13423604
                    picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 90)-(UserControl.Width - 510, UserControl.Extender.Top + 300), 13423604
                    picForm.Line (UserControl.Width - 120, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 300), 13423604
                End If
                If bCloseButtonHover = True And bCloseButtonClicked = True Then
                    picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 90), 13423604
                    picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 300)-(UserControl.Width - 105, UserControl.Extender.Top + 300), 13423604
                    picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 90)-(UserControl.Width - 510, UserControl.Extender.Top + 300), 13423604
                    picForm.Line (UserControl.Width - 120, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 300), 13423604
                End If
                
                ' Close Button Top Gradient Base Colour
                iVertical = 105
                For I = 1 To 6
                    If bCloseButtonHover = False And bCloseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTop
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTopHover
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTopClicked
                    End If
                    iVertical = iVertical + 15
                Next
                
                ' Button Bottom Gradient Base Colour
                For I = 1 To 7
                    If bCloseButtonHover = False And bCloseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottom(I)
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottomHover(I)
                    ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottomClicked(I)
                    End If
                    iVertical = iVertical + 15
                Next
                
                ' Draw the Close Button Display
                ' Outside Borders
                picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 135)-(UserControl.Width - 330, UserControl.Extender.Top + 135), &H665653
                picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 135)-(UserControl.Width - 240, UserControl.Extender.Top + 135), &H665653
                
                picForm.Line (UserControl.Width - 390, UserControl.Extender.Top + 150)-(UserControl.Width - 375, UserControl.Extender.Top + 150), &H665653
                picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 150)-(UserControl.Width - 315, UserControl.Extender.Top + 150), &H665653
                picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 150)-(UserControl.Width - 285, UserControl.Extender.Top + 150), &H665653
                picForm.Line (UserControl.Width - 240, UserControl.Extender.Top + 150)-(UserControl.Width - 225, UserControl.Extender.Top + 150), &H665653
                
                picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 165)-(UserControl.Width - 360, UserControl.Extender.Top + 165), &H665653
                picForm.Line (UserControl.Width - 315, UserControl.Extender.Top + 165)-(UserControl.Width - 300, UserControl.Extender.Top + 165), &H665653
                picForm.Line (UserControl.Width - 255, UserControl.Extender.Top + 165)-(UserControl.Width - 240, UserControl.Extender.Top + 165), &H665653
               
                picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 180)-(UserControl.Width - 345, UserControl.Extender.Top + 180), &H665653
                picForm.Line (UserControl.Width - 270, UserControl.Extender.Top + 180)-(UserControl.Width - 255, UserControl.Extender.Top + 180), &H665653
        
                picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 195)-(UserControl.Width - 330, UserControl.Extender.Top + 195), &H665653
                picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 195)-(UserControl.Width - 270, UserControl.Extender.Top + 195), &H665653
        
                picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 210)-(UserControl.Width - 345, UserControl.Extender.Top + 210), &H665653
                picForm.Line (UserControl.Width - 270, UserControl.Extender.Top + 210)-(UserControl.Width - 255, UserControl.Extender.Top + 210), &H665653
        
                picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 225)-(UserControl.Width - 360, UserControl.Extender.Top + 225), &H665653
                picForm.Line (UserControl.Width - 315, UserControl.Extender.Top + 225)-(UserControl.Width - 300, UserControl.Extender.Top + 225), &H665653
                picForm.Line (UserControl.Width - 255, UserControl.Extender.Top + 225)-(UserControl.Width - 240, UserControl.Extender.Top + 225), &H665653
        
                picForm.Line (UserControl.Width - 390, UserControl.Extender.Top + 240)-(UserControl.Width - 375, UserControl.Extender.Top + 240), &H665653
                picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 240)-(UserControl.Width - 315, UserControl.Extender.Top + 240), &H665653
                picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 240)-(UserControl.Width - 285, UserControl.Extender.Top + 240), &H665653
                picForm.Line (UserControl.Width - 240, UserControl.Extender.Top + 240)-(UserControl.Width - 225, UserControl.Extender.Top + 240), &H665653
        
                picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 255)-(UserControl.Width - 330, UserControl.Extender.Top + 255), &H665653
                picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 255)-(UserControl.Width - 240, UserControl.Extender.Top + 255), &H665653
        
                ' Inside Button Colours
                picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 150)-(UserControl.Width - 330, UserControl.Extender.Top + 150), &HFFFFFF
                picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 150)-(UserControl.Width - 240, UserControl.Extender.Top + 150), &HFFFFFF
                
                picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 165)-(UserControl.Width - 315, UserControl.Extender.Top + 165), &HFFFFFF
                picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 165)-(UserControl.Width - 255, UserControl.Extender.Top + 165), &HFFFFFF
                
                picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 180)-(UserControl.Width - 270, UserControl.Extender.Top + 180), &HFFFFFF
                
                picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 195)-(UserControl.Width - 285, UserControl.Extender.Top + 195), &HE9E9E9
                
                picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 210)-(UserControl.Width - 270, UserControl.Extender.Top + 210), &HE2E2E2
                
                picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 225)-(UserControl.Width - 315, UserControl.Extender.Top + 225), &HDCDCDC
                picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 225)-(UserControl.Width - 255, UserControl.Extender.Top + 225), &HDCDCDC
                
                picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 240)-(UserControl.Width - 330, UserControl.Extender.Top + 240), &HD7D7D7
                picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 240)-(UserControl.Width - 240, UserControl.Extender.Top + 240), &HD7D7D7
            End If
        End If
    End If
Errhandler:
End Sub

Private Sub DrawMinimiseButton()
On Error GoTo Errhandler
Dim Vi As Byte
Dim VWidth As Long
Dim VTop As Long

VWidth = UserControl.Width
VTop = UserControl.Extender.Top
    
    
    If Style_Type = Vista_Aero Then
            ' Draw the Buttons
            If bMinimiseButton = True Then
                ' Draw the Minimise Buttons Outside Border
                
                If EnableMinimiseButton = True Then
                    picForm.Line (VWidth - 1470, VTop + 15)-(VWidth - 1050, VTop + 15), lButtonOuterBorder
                    picForm.Line (VWidth - 1425, VTop + 255)-(VWidth - 1050, VTop + 255), lButtonOuterBorder
                    picForm.Line (VWidth - 1485, VTop + 30)-(VWidth - 1485, VTop + 210), lButtonOuterBorder
                    picForm.Line (VWidth - 1050, VTop + 30)-(VWidth - 1050, VTop + 255), lButtonOuterBorder
            
                    For Vi = 15 To 60 Step 15
                        picForm.Line (VWidth - 1500 + Vi, VTop + 190 + Vi)-(VWidth - 1515 + Vi, VTop + 205 + Vi), lButtonOuterBorder
                    Next
                    
                    ' Draw the Buttons Inside Border
                    picForm.Line (VWidth - 1470, VTop + 30)-(VWidth - 1070, VTop + 30), lButtonInnerBorder
                    picForm.Line (VWidth - 1425, VTop + 240)-(VWidth - 1055, VTop + 240), lButtonInnerBorder
                    picForm.Line (VWidth - 1070, VTop + 90 - 60)-(VWidth - 1070, VTop + 240), lButtonInnerBorder
                    picForm.Line (VWidth - 1470, VTop + 90 - 60)-(VWidth - 1470, VTop + 210), lButtonInnerBorder
            
                    For Vi = 15 To 45 Step 15
                        picForm.Line (VWidth - 1485 + Vi, VTop + 190 + Vi)-(VWidth - 1500 + Vi, VTop + 205 + Vi), lButtonInnerBorder
                    Next
            
                    ' Button Top Gradient Base Colour
                    iVertical = 105
                    For I = 1 To 6
                        If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                            picForm.Line (VWidth - 1455, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), lButtonGradientTop
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                            picForm.Line (VWidth - 1455, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), lButtonGradientTopHover
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                            picForm.Line (VWidth - 1455, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), lButtonGradientTopClicked
                        End If
                        iVertical = iVertical + 15
                    Next
                    
                    ' Button Bottom Gradient Base Colour
                    Vi = 0
                    For I = 1 To 7
                        If I >= 7 Then
                            Vi = Vi + 15
                        End If
                        
                        If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                            picForm.Line (VWidth - 1455 + Vi, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), lButtonGradientBottom(I)
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                            picForm.Line (VWidth - 1455 + Vi, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), lButtonGradientBottomHover(I)
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                            picForm.Line (VWidth - 1455 + Vi, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), lButtonGradientBottomClicked(I)
                        End If
                        iVertical = iVertical + 15
                    Next
                     
                    ' Draw the Minimise Button Display
                    ' Outside Borders
            
                    picForm.Line (VWidth - 1335, VTop + 195 - 60)-(VWidth - 1190, VTop + 195 - 60), &H665653 ' Top Border
                    picForm.Line (VWidth - 1335, VTop + 255 - 60)-(VWidth - 1190, VTop + 255 - 60), &H665653 ' Bottom Border
                    picForm.Line (VWidth - 1350, VTop + 210 - 60)-(VWidth - 1350, VTop + 255 - 60), &H665653 ' Left Border
                    picForm.Line (VWidth - 1190, VTop + 210 - 60)-(VWidth - 1190, VTop + 255 - 60), &H665653 ' Right Border
                
                    ' Inside Button Display
                    picForm.Line (VWidth - 1335, VTop + 210 - 60)-(VWidth - 1190, VTop + 210 - 60), &HFFFFFF ' Top Border
                    picForm.Line (VWidth - 1335, VTop + 225 - 60)-(VWidth - 1190, VTop + 225 - 60), &HDCDCDC ' Middle Border
                    picForm.Line (VWidth - 1335, VTop + 240 - 60)-(VWidth - 1190, VTop + 240 - 60), &HD7D7D7 ' Bottom Border
                Else
                    'CreateGradientButom(35, 70, lButtonGradientBottom(1))
                    picForm.Line (VWidth - 1470, VTop + 15)-(VWidth - 1050, VTop + 15), CreateGradientButom(35, 70, lButtonOuterBorder)
                    picForm.Line (VWidth - 1425, VTop + 255)-(VWidth - 1050, VTop + 255), CreateGradientButom(35, 70, lButtonOuterBorder)
                    picForm.Line (VWidth - 1485, VTop + 30)-(VWidth - 1485, VTop + 210), CreateGradientButom(35, 70, lButtonOuterBorder)
                    'picForm.Line (VWidth - 1050, VTop + 30)-(VWidth - 1050, VTop + 255), CreateGradientButom(35, 70, lButtonOuterBorder)
            
                    For Vi = 15 To 60 Step 15
                        picForm.Line (VWidth - 1500 + Vi, VTop + 190 + Vi)-(VWidth - 1515 + Vi, VTop + 205 + Vi), CreateGradientButom(35, 70, lButtonOuterBorder)
                    Next
                    
                    ' Draw the Buttons Inside Border
                    picForm.Line (VWidth - 1470, VTop + 30)-(VWidth - 1070, VTop + 30), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (VWidth - 1425, VTop + 240)-(VWidth - 1055, VTop + 240), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (VWidth - 1070, VTop + 90 - 60)-(VWidth - 1070, VTop + 240), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (VWidth - 1470, VTop + 90 - 60)-(VWidth - 1470, VTop + 210), CreateGradientButom(35, 70, lButtonInnerBorder)
            
                    For Vi = 15 To 45 Step 15
                        picForm.Line (VWidth - 1485 + Vi, VTop + 190 + Vi)-(VWidth - 1500 + Vi, VTop + 205 + Vi), CreateGradientButom(35, 70, lButtonInnerBorder)
                    Next
            
                    ' Button Top Gradient Base Colour
                    iVertical = 105
                    For I = 1 To 6
                        picForm.Line (VWidth - 1455, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), CreateGradientButom(35, 70, lButtonGradientTop)
                        iVertical = iVertical + 15
                    Next
                    
                    ' Button Bottom Gradient Base Colour
                    Vi = 0
                    For I = 1 To 7
                        If I >= 7 Then
                            Vi = Vi + 15
                        End If
                        picForm.Line (VWidth - 1455 + Vi, VTop + iVertical - 60)-(VWidth - 1070, VTop + iVertical - 60), CreateGradientButom(35, 70, lButtonGradientBottom(I))
                        iVertical = iVertical + 15
                    Next
                     
            
                    picForm.Line (VWidth - 1335, VTop + 195 - 60)-(VWidth - 1190, VTop + 195 - 60), CreateGradientButom(35, 70, &H665653)  ' Top Border
                    picForm.Line (VWidth - 1335, VTop + 255 - 60)-(VWidth - 1190, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Bottom Border
                    picForm.Line (VWidth - 1350, VTop + 210 - 60)-(VWidth - 1350, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Left Border
                    picForm.Line (VWidth - 1190, VTop + 210 - 60)-(VWidth - 1190, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Right Border
                
                    ' Inside Button Display
                    picForm.Line (VWidth - 1335, VTop + 210 - 60)-(VWidth - 1190, VTop + 210 - 60), CreateGradientButom(35, 70, &HFFFFFF) ' Top Border
                    picForm.Line (VWidth - 1335, VTop + 225 - 60)-(VWidth - 1190, VTop + 225 - 60), CreateGradientButom(35, 70, &HDCDCDC)  ' Middle Border
                    picForm.Line (VWidth - 1335, VTop + 240 - 60)-(VWidth - 1190, VTop + 240 - 60), CreateGradientButom(35, 70, &HD7D7D7)  ' Bottom Border
                End If
            End If
    Else
            ' Draw the Buttons
            If bMinimiseButton = True Then
                If bEnableMinimiseButton = True Then
                    ' Draw the Minimise Buttons Outside Border
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 75)-(UserControl.Width - 1050, UserControl.Extender.Top + 75), lButtonOuterBorder
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 315)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), lButtonOuterBorder
                    picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 90)-(UserControl.Width - 1485, UserControl.Extender.Top + 315), lButtonOuterBorder
                    picForm.Line (UserControl.Width - 1050, UserControl.Extender.Top + 90)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), lButtonOuterBorder
            
                    ' Draw the Buttons Inside Border
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 90), lButtonInnerBorder
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 300)-(UserControl.Width - 1055, UserControl.Extender.Top + 300), lButtonInnerBorder
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1470, UserControl.Extender.Top + 300), lButtonInnerBorder
                    picForm.Line (UserControl.Width - 1070, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 300), lButtonInnerBorder
            
                    ' Button Top Gradient Base Colour
                    iVertical = 105
                    For I = 1 To 6
                        If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                            picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTop
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                            picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTopHover
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                            picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTopClicked
                        End If
                        iVertical = iVertical + 15
                    Next
                    
                    ' Button Bottom Gradient Base Colour
                    For I = 1 To 7
                        If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                            picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottom(I)
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                            picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottomHover(I)
                        ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                            picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottomClicked(I)
                        End If
                        iVertical = iVertical + 15
                    Next
                     
                    ' Draw the Minimise Button Display
                    ' Outside Borders
            
                    picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 195)-(UserControl.Width - 1190, UserControl.Extender.Top + 195), &H665653    ' Top Border
                    picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 255)-(UserControl.Width - 1190, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
                    picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 210)-(UserControl.Width - 1350, UserControl.Extender.Top + 255), &H665653    ' Left Border
                    picForm.Line (UserControl.Width - 1190, UserControl.Extender.Top + 210)-(UserControl.Width - 1190, UserControl.Extender.Top + 255), &H665653    ' Right Border
                
                    ' Inside Button Display
                    picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 210)-(UserControl.Width - 1190, UserControl.Extender.Top + 210), &HFFFFFF    ' Top Border
                    picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 225)-(UserControl.Width - 1190, UserControl.Extender.Top + 225), &HDCDCDC    ' Middle Border
                    picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 240)-(UserControl.Width - 1190, UserControl.Extender.Top + 240), &HD7D7D7    ' Bottom Border
                Else
                    ' Draw the Minimise Buttons Outside Border
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 75)-(UserControl.Width - 1050, UserControl.Extender.Top + 75), CreateGradientButom(35, 70, lButtonOuterBorder)
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 315)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
                    picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 90)-(UserControl.Width - 1485, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
                    picForm.Line (UserControl.Width - 1050, UserControl.Extender.Top + 90)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
            
                    ' Draw the Buttons Inside Border
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 90), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 300)-(UserControl.Width - 1055, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1470, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (UserControl.Width - 1070, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
            
                    ' Button Top Gradient Base Colour
                    iVertical = 105
                    For I = 1 To 6
                        picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), CreateGradientButom(35, 70, lButtonGradientTop)
                        iVertical = iVertical + 15
                    Next
                    
                    ' Button Bottom Gradient Base Colour
                    For I = 1 To 7
                        picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), CreateGradientButom(35, 70, lButtonGradientBottom(I))
                        iVertical = iVertical + 15
                    Next
                     
                    ' Draw the Minimise Button Display
                    ' Outside Borders
                    picForm.Line (VWidth - 1335, VTop + 195)-(VWidth - 1190, VTop + 195), CreateGradientButom(35, 70, &H665653)   ' Top Border
                    picForm.Line (VWidth - 1335, VTop + 255)-(VWidth - 1190, VTop + 255), CreateGradientButom(35, 70, &H665653)   ' Bottom Border
                    picForm.Line (VWidth - 1350, VTop + 210)-(VWidth - 1350, VTop + 255), CreateGradientButom(35, 70, &H665653)   ' Left Border
                    picForm.Line (VWidth - 1190, VTop + 210)-(VWidth - 1190, VTop + 255), CreateGradientButom(35, 70, &H665653)   ' Right Border
                
                    ' Inside Button Display
                    picForm.Line (VWidth - 1335, VTop + 210)-(VWidth - 1190, VTop + 210), CreateGradientButom(35, 70, &HFFFFFF)   ' Top Border
                    picForm.Line (VWidth - 1335, VTop + 225)-(VWidth - 1190, VTop + 225), CreateGradientButom(35, 70, &HDCDCDC)   ' Middle Border
                    picForm.Line (VWidth - 1335, VTop + 240)-(VWidth - 1190, VTop + 240), CreateGradientButom(35, 70, &HD7D7D7)   ' Bottom Border
                End If
            End If
    End If
Errhandler:
End Sub

Private Sub DrawMaximiseButton()
On Error GoTo Errhandler
Dim VT As Byte
Dim VWidth As Long
Dim VTop As Long

VT = 45
VWidth = UserControl.Width
VTop = UserControl.Extender.Top

    If Style_Type = Vista_Aero Then
            If bMaximiseButton = True Then
                If bEnableMaximiseButton Then
                        ' Draw the Maximise Buttons Outside Border
                        picForm.Line (VWidth - 990 - VT, VTop + 75 - 60)-(VWidth - 570 - VT, VTop + 75 - 60), lButtonOuterBorder
                        picForm.Line (VWidth - 990 - VT, VTop + 315 - 60)-(VWidth - 570 - VT, VTop + 315 - 60), lButtonOuterBorder
                        picForm.Line (VWidth - 1005 - VT, VTop + 90 - 60)-(VWidth - 1005 - VT, VTop + 315 - 60), lButtonOuterBorder
                        picForm.Line (VWidth - 570 - VT, VTop + 90 - 60)-(VWidth - 570 - VT, VTop + 315 - 60), lButtonOuterBorder
                
                        ' Draw the Buttons Inside Border
                        picForm.Line (VWidth - 980 - VT, VTop + 90 - 60)-(VWidth - 590 - VT, VTop + 90 - 60), lButtonInnerBorder
                        picForm.Line (VWidth - 990 - VT, VTop + 300 - 60)-(VWidth - 575 - VT, VTop + 300 - 60), lButtonInnerBorder
                        picForm.Line (VWidth - 990 - VT, VTop + 90 - 60)-(VWidth - 990 - VT, VTop + 300 - 60), lButtonInnerBorder
                        picForm.Line (VWidth - 590 - VT, VTop + 90 - 60)-(VWidth - 590 - VT, VTop + 300 - 60), lButtonInnerBorder
                
                        ' Button Top Gradient Base Colour
                        iVertical = 105
                        For I = 1 To 6
                            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                                picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), lButtonGradientTop
                            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                                picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), lButtonGradientTopHover
                            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                                picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), lButtonGradientTopClicked
                            End If
                            iVertical = iVertical + 15
                        Next
                
                        ' Button Bottom Gradient Base Colour
                        For I = 1 To 7
                            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                                picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), lButtonGradientBottom(I)
                            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                                picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), lButtonGradientBottomHover(I)
                            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                                picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), lButtonGradientBottomClicked(I)
                            End If
                            iVertical = iVertical + 15
                        Next
                        
                        ' Draw the Maximise Button Display
                        ' Inside Button Display
                        If myForm.Tag <> "vbMaximized" Then
                            picForm.Line (VWidth - 855 - VT, VTop + 150 - 60)-(VWidth - 705 - VT, VTop + 150 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 165 - 60)-(VWidth - 705 - VT, VTop + 165 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 180 - 60)-(VWidth - 705 - VT, VTop + 180 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 195 - 60)-(VWidth - 705 - VT, VTop + 195 - 60), &HE9E9E9
                            picForm.Line (VWidth - 855 - VT, VTop + 210 - 60)-(VWidth - 705 - VT, VTop + 210 - 60), &HE2E2E2
                            picForm.Line (VWidth - 855 - VT, VTop + 225 - 60)-(VWidth - 705 - VT, VTop + 225 - 60), &HDCDCDC
                            picForm.Line (VWidth - 855 - VT, VTop + 240 - 60)-(VWidth - 705 - VT, VTop + 240 - 60), &HD7D7D7
                            
                            ' Outside Borders
                            
                            picForm.Line (VWidth - 855 - VT, VTop + 135 - 60)-(VWidth - 705 - VT, VTop + 135 - 60), &H665653 ' Top Border
                            picForm.Line (VWidth - 855 - VT, VTop + 255 - 60)-(VWidth - 705 - VT, VTop + 255 - 60), &H665653 ' Bottom Border
                            picForm.Line (VWidth - 870 - VT, VTop + 150 - 60)-(VWidth - 870 - VT, VTop + 255 - 60), &H665653 ' Left Border
                            picForm.Line (VWidth - 705 - VT, VTop + 150 - 60)-(VWidth - 705 - VT, VTop + 255 - 60), &H665653 ' Right Border
                    
                            picForm.Line (VWidth - 825 - VT, VTop + 180 - 60)-(VWidth - 735 - VT, VTop + 180 - 60), &H665653 ' Top Border
                            picForm.Line (VWidth - 825 - VT, VTop + 215 - 60)-(VWidth - 735 - VT, VTop + 215 - 60), &H665653 ' Bottom Border
                            picForm.Line (VWidth - 825 - VT, VTop + 180 - 60)-(VWidth - 825 - VT, VTop + 215 - 60), &H665653 ' Left Border
                            picForm.Line (VWidth - 750 - VT, VTop + 180 - 60)-(VWidth - 750 - VT, VTop + 215 - 60), &H665653 ' Right Border
                            
                            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                                picForm.Line (VWidth - 810 - VT, VTop + 195 - 60)-(VWidth - 750 - VT, VTop + 195 - 60), lButtonGradientBottom(1)
                            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                                picForm.Line (VWidth - 810 - VT, VTop + 195 - 60)-(VWidth - 750 - VT, VTop + 195 - 60), lButtonGradientBottomHover(1)
                            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                                picForm.Line (VWidth - 810 - VT, VTop + 195 - 60)-(VWidth - 750 - VT, VTop + 195 - 60), lButtonGradientBottomClicked(1)
                            End If
                            
                            ' Outside Borders
                        Else
                            picForm.Line (VWidth - 810 - VT, VTop + 135 - 60)-(VWidth - 705 - VT, VTop + 135 - 60), &HFFFFFF
                            picForm.Line (VWidth - 810 - VT, VTop + 150 - 60)-(VWidth - 795 - VT, VTop + 150 - 60), &HFFFFFF
                            picForm.Line (VWidth - 795 - VT, VTop + 150 - 60)-(VWidth - 720 - VT, VTop + 150 - 60), lButtonGradientBottom(1)
                            picForm.Line (VWidth - 720 - VT, VTop + 150 - 60)-(VWidth - 705 - VT, VTop + 150 - 60), &HFFFFFF
                            picForm.Line (VWidth - 720 - VT, VTop + 165 - 60)-(VWidth - 705 - VT, VTop + 165 - 60), &HFFFFFF
                            picForm.Line (VWidth - 720 - VT, VTop + 180 - 60)-(VWidth - 705 - VT, VTop + 180 - 60), &HFFFFFF
                            picForm.Line (VWidth - 735 - VT, VTop + 195 - 60)-(VWidth - 705 - VT, VTop + 195 - 60), &HFFFFFF
                            picForm.Line (VWidth - 735 - VT, VTop + 150 - 60)-(VWidth - 735 - VT, VTop + 195 - 60), lButtonGradientBottom(1)
                            
                            picForm.Line (VWidth - 810 - VT, VTop + 120 - 60)-(VWidth - 705 - VT, VTop + 120 - 60), &H665653 ' Top Border
                            picForm.Line (VWidth - 735 - VT, VTop + 210 - 60)-(VWidth - 705 - VT, VTop + 210 - 60), &H665653 ' Bottom Border
                            picForm.Line (VWidth - 825 - VT, VTop + 120 - 60)-(VWidth - 825 - VT, VTop + 170 - 60), &H665653 ' Left Border
                            picForm.Line (VWidth - 705 - VT, VTop + 120 - 60)-(VWidth - 705 - VT, VTop + 225 - 60), &H665653 ' Right Border
                
                            picForm.Line (VWidth - 855 - VT, VTop + 180 - 60)-(VWidth - 750 - VT, VTop + 180 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 195 - 60)-(VWidth - 840 - VT, VTop + 195 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 195 - 60)-(VWidth - 840 - VT, VTop + 195 - 60), &HFFFFFF
                            picForm.Line (VWidth - 765 - VT, VTop + 195 - 60)-(VWidth - 750 - VT, VTop + 195 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 210 - 60)-(VWidth - 840 - VT, VTop + 210 - 60), &HFFFFFF
                            picForm.Line (VWidth - 765 - VT, VTop + 210 - 60)-(VWidth - 750 - VT, VTop + 210 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 225 - 60)-(VWidth - 840 - VT, VTop + 225 - 60), &HFFFFFF
                            picForm.Line (VWidth - 765 - VT, VTop + 225 - 60)-(VWidth - 750 - VT, VTop + 225 - 60), &HFFFFFF
                            picForm.Line (VWidth - 855 - VT, VTop + 240 - 60)-(VWidth - 750 - VT, VTop + 240 - 60), &HFFFFFF
                
                            picForm.Line (VWidth - 855 - VT, VTop + 165 - 60)-(VWidth - 750 - VT, VTop + 165 - 60), &H665653 ' Top Border
                            picForm.Line (VWidth - 855 - VT, VTop + 255 - 60)-(VWidth - 750 - VT, VTop + 255 - 60), &H665653 ' Bottom Border
                            picForm.Line (VWidth - 870 - VT, VTop + 165 - 60)-(VWidth - 870 - VT, VTop + 270 - 60), &H665653 ' Left Border
                            picForm.Line (VWidth - 750 - VT, VTop + 165 - 60)-(VWidth - 750 - VT, VTop + 270 - 60), &H665653 ' Right Border
                        End If
                Else
                    'Draw the Maximise Buttons Outside Border
                    picForm.Line (VWidth - 990 - VT, VTop + 75 - 60)-(VWidth - 570 - VT, VTop + 75 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
                    picForm.Line (VWidth - 990 - VT, VTop + 315 - 60)-(VWidth - 570 - VT, VTop + 315 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
                    'picForm.Line (VWidth - 1005 - VT, VTop + 90 - 60)-(VWidth - 1005 - VT, VTop + 315 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
                    'picForm.Line (VWidth - 570 - VT, VTop + 90 - 60)-(VWidth - 570 - VT, VTop + 315 - 60), CreateGradientButom(35, 70, lButtonOuterBorder)
            
                    'Draw the Buttons Inside Border
                    picForm.Line (VWidth - 980 - VT, VTop + 90 - 60)-(VWidth - 590 - VT, VTop + 90 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (VWidth - 990 - VT, VTop + 300 - 60)-(VWidth - 575 - VT, VTop + 300 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (VWidth - 990 - VT, VTop + 90 - 60)-(VWidth - 990 - VT, VTop + 300 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
                    picForm.Line (VWidth - 590 - VT, VTop + 90 - 60)-(VWidth - 590 - VT, VTop + 300 - 60), CreateGradientButom(35, 70, lButtonInnerBorder)
            
                    'Button Top Gradient Base Colour
                    iVertical = 105
                    For I = 1 To 6
                        picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), CreateGradientButom(35, 70, lButtonGradientTop)
                        iVertical = iVertical + 15
                    Next
            
                    'Button Bottom Gradient Base Colour
                    For I = 1 To 7
                        picForm.Line (VWidth - 980 - VT, VTop + iVertical - 60)-(VWidth - 590 - VT, VTop + iVertical - 60), CreateGradientButom(35, 70, lButtonGradientBottom(I))
                        iVertical = iVertical + 15
                    Next
                    
                    'Draw the Maximise Button Display
                    'Inside Button Display
                    If myForm.Tag <> "vbMaximized" Then
                        picForm.Line (VWidth - 855 - VT, VTop + 150 - 60)-(VWidth - 705 - VT, VTop + 150 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 165 - 60)-(VWidth - 705 - VT, VTop + 165 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 180 - 60)-(VWidth - 705 - VT, VTop + 180 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 195 - 60)-(VWidth - 705 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, &HE9E9E9)
                        picForm.Line (VWidth - 855 - VT, VTop + 210 - 60)-(VWidth - 705 - VT, VTop + 210 - 60), CreateGradientButom(35, 70, &HE2E2E2)
                        picForm.Line (VWidth - 855 - VT, VTop + 225 - 60)-(VWidth - 705 - VT, VTop + 225 - 60), CreateGradientButom(35, 70, &HDCDCDC)
                        picForm.Line (VWidth - 855 - VT, VTop + 240 - 60)-(VWidth - 705 - VT, VTop + 240 - 60), CreateGradientButom(35, 70, &HD7D7D7)
                        
                        'Outside Borders
                        picForm.Line (VWidth - 855 - VT, VTop + 135 - 60)-(VWidth - 705 - VT, VTop + 135 - 60), CreateGradientButom(35, 70, &H665653) ' Top Border
                        picForm.Line (VWidth - 855 - VT, VTop + 255 - 60)-(VWidth - 705 - VT, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Bottom Border
                        picForm.Line (VWidth - 870 - VT, VTop + 150 - 60)-(VWidth - 870 - VT, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Left Border
                        picForm.Line (VWidth - 705 - VT, VTop + 150 - 60)-(VWidth - 705 - VT, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Right Border
                
                        picForm.Line (VWidth - 825 - VT, VTop + 180 - 60)-(VWidth - 735 - VT, VTop + 180 - 60), CreateGradientButom(35, 70, &H665653) ' Top Border
                        picForm.Line (VWidth - 825 - VT, VTop + 215 - 60)-(VWidth - 735 - VT, VTop + 215 - 60), CreateGradientButom(35, 70, &H665653) ' Bottom Border
                        picForm.Line (VWidth - 825 - VT, VTop + 180 - 60)-(VWidth - 825 - VT, VTop + 215 - 60), CreateGradientButom(35, 70, &H665653) ' Left Border
                        picForm.Line (VWidth - 750 - VT, VTop + 180 - 60)-(VWidth - 750 - VT, VTop + 215 - 60), CreateGradientButom(35, 70, &H665653) ' Right Border
                        
                        picForm.Line (VWidth - 810 - VT, VTop + 195 - 60)-(VWidth - 750 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, lButtonGradientBottom(1))
                        
                        'Outside Borders
                    Else
                        picForm.Line (VWidth - 810 - VT, VTop + 135 - 60)-(VWidth - 705 - VT, VTop + 135 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 810 - VT, VTop + 150 - 60)-(VWidth - 795 - VT, VTop + 150 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 795 - VT, VTop + 150 - 60)-(VWidth - 720 - VT, VTop + 150 - 60), CreateGradientButom(35, 70, lButtonGradientBottom(1))
                        picForm.Line (VWidth - 720 - VT, VTop + 150 - 60)-(VWidth - 705 - VT, VTop + 150 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 720 - VT, VTop + 165 - 60)-(VWidth - 705 - VT, VTop + 165 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 720 - VT, VTop + 180 - 60)-(VWidth - 705 - VT, VTop + 180 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 735 - VT, VTop + 195 - 60)-(VWidth - 705 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 735 - VT, VTop + 150 - 60)-(VWidth - 735 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, lButtonGradientBottom(1))
                        
                        picForm.Line (VWidth - 810 - VT, VTop + 120 - 60)-(VWidth - 705 - VT, VTop + 120 - 60), CreateGradientButom(35, 70, &H665653) ' Top Border
                        picForm.Line (VWidth - 735 - VT, VTop + 210 - 60)-(VWidth - 705 - VT, VTop + 210 - 60), CreateGradientButom(35, 70, &H665653) ' Bottom Border
                        picForm.Line (VWidth - 825 - VT, VTop + 120 - 60)-(VWidth - 825 - VT, VTop + 170 - 60), CreateGradientButom(35, 70, &H665653) ' Left Border
                        picForm.Line (VWidth - 705 - VT, VTop + 120 - 60)-(VWidth - 705 - VT, VTop + 225 - 60), CreateGradientButom(35, 70, &H665653) ' Right Border
            
                        picForm.Line (VWidth - 855 - VT, VTop + 180 - 60)-(VWidth - 750 - VT, VTop + 180 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 195 - 60)-(VWidth - 840 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 195 - 60)-(VWidth - 840 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 765 - VT, VTop + 195 - 60)-(VWidth - 750 - VT, VTop + 195 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 210 - 60)-(VWidth - 840 - VT, VTop + 210 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 765 - VT, VTop + 210 - 60)-(VWidth - 750 - VT, VTop + 210 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 225 - 60)-(VWidth - 840 - VT, VTop + 225 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 765 - VT, VTop + 225 - 60)-(VWidth - 750 - VT, VTop + 225 - 60), CreateGradientButom(35, 70, &HFFFFFF)
                        picForm.Line (VWidth - 855 - VT, VTop + 240 - 60)-(VWidth - 750 - VT, VTop + 240 - 60), CreateGradientButom(35, 70, &HFFFFFF)
            
                        picForm.Line (VWidth - 855 - VT, VTop + 165 - 60)-(VWidth - 750 - VT, VTop + 165 - 60), CreateGradientButom(35, 70, &H665653) ' Top Border
                        picForm.Line (VWidth - 855 - VT, VTop + 255 - 60)-(VWidth - 750 - VT, VTop + 255 - 60), CreateGradientButom(35, 70, &H665653) ' Bottom Border
                        picForm.Line (VWidth - 870 - VT, VTop + 165 - 60)-(VWidth - 870 - VT, VTop + 270 - 60), CreateGradientButom(35, 70, &H665653) ' Left Border
                        picForm.Line (VWidth - 750 - VT, VTop + 165 - 60)-(VWidth - 750 - VT, VTop + 270 - 60), CreateGradientButom(35, 70, &H665653) ' Right Border
                    End If
                End If
            End If
    Else
        If bMaximiseButton = True Then
            If bEnableMaximiseButton Then
                ' Draw the Maximise Buttons Outside Border
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 75)-(UserControl.Width - 570, UserControl.Extender.Top + 75), lButtonOuterBorder
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 315)-(UserControl.Width - 570, UserControl.Extender.Top + 315), lButtonOuterBorder
                picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 90)-(UserControl.Width - 1005, UserControl.Extender.Top + 315), lButtonOuterBorder
                picForm.Line (UserControl.Width - 570, UserControl.Extender.Top + 90)-(UserControl.Width - 570, UserControl.Extender.Top + 315), lButtonOuterBorder
        
                ' Draw the Buttons Inside Border
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 90), lButtonInnerBorder
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 300)-(UserControl.Width - 575, UserControl.Extender.Top + 300), lButtonInnerBorder
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 90)-(UserControl.Width - 990, UserControl.Extender.Top + 300), lButtonInnerBorder
                picForm.Line (UserControl.Width - 590, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 300), lButtonInnerBorder
        
                ' Button Top Gradient Base Colour
                iVertical = 105
                For I = 1 To 6
                    If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTop
                    ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTopHover
                    ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTopClicked
                    End If
                    iVertical = iVertical + 15
                Next
        
                ' Button Bottom Gradient Base Colour
                For I = 1 To 7
                    If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottom(I)
                    ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottomHover(I)
                    ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottomClicked(I)
                    End If
                    iVertical = iVertical + 15
                Next
                
                ' Draw the Maximise Button Display
                ' Inside Button Display
                If myForm.Tag <> "vbMaximized" Then
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), &HE9E9E9
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), &HE2E2E2
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 705, UserControl.Extender.Top + 225), &HDCDCDC
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 705, UserControl.Extender.Top + 240), &HD7D7D7
                    
                    ' Outside Borders
                    
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), &H665653    ' Top Border
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 705, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
                    picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 150)-(UserControl.Width - 870, UserControl.Extender.Top + 255), &H665653    ' Left Border
                    picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 255), &H665653    ' Right Border
            
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 735, UserControl.Extender.Top + 180), &H665653    ' Top Border
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 215)-(UserControl.Width - 735, UserControl.Extender.Top + 215), &H665653    ' Bottom Border
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 825, UserControl.Extender.Top + 215), &H665653    ' Left Border
                    picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 215), &H665653    ' Right Border
                    
                    If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottom(1)
                    ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                        picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottomHover(1)
                    ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                        picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottomClicked(1)
                    End If
                    
                    ' Outside Borders
                Else
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), &HFFFFFF
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 150)-(UserControl.Width - 795, UserControl.Extender.Top + 150), &HFFFFFF
                    picForm.Line (UserControl.Width - 795, UserControl.Extender.Top + 150)-(UserControl.Width - 720, UserControl.Extender.Top + 150), lButtonGradientBottom(1)
                    picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), &HFFFFFF
                    picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), &HFFFFFF
                    picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), &HFFFFFF
                    picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), &HFFFFFF
                    picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 150)-(UserControl.Width - 735, UserControl.Extender.Top + 195), lButtonGradientBottom(1)
                    
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 120)-(UserControl.Width - 705, UserControl.Extender.Top + 120), &H665653    ' Top Border
                    picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), &H665653    ' Bottom Border
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 120)-(UserControl.Width - 825, UserControl.Extender.Top + 170), &H665653    ' Left Border
                    picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 120)-(UserControl.Width - 705, UserControl.Extender.Top + 225), &H665653    ' Right Border
        
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 180), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 840, UserControl.Extender.Top + 195), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 840, UserControl.Extender.Top + 195), &HFFFFFF
                    picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 840, UserControl.Extender.Top + 210), &HFFFFFF
                    picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 210)-(UserControl.Width - 750, UserControl.Extender.Top + 210), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 840, UserControl.Extender.Top + 225), &HFFFFFF
                    picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 225)-(UserControl.Width - 750, UserControl.Extender.Top + 225), &HFFFFFF
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 750, UserControl.Extender.Top + 240), &HFFFFFF
        
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 750, UserControl.Extender.Top + 165), &H665653    ' Top Border
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 750, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
                    picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 165)-(UserControl.Width - 870, UserControl.Extender.Top + 270), &H665653    ' Left Border
                    picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 165)-(UserControl.Width - 750, UserControl.Extender.Top + 270), &H665653    ' Right Border
                End If
            Else
                'Draw the Maximise Buttons Outside Border
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 75)-(UserControl.Width - 570, UserControl.Extender.Top + 75), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 315)-(UserControl.Width - 570, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 90)-(UserControl.Width - 1005, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
                picForm.Line (UserControl.Width - 570, UserControl.Extender.Top + 90)-(UserControl.Width - 570, UserControl.Extender.Top + 315), CreateGradientButom(35, 70, lButtonOuterBorder)
        
                ' Draw the Buttons Inside Border
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 90), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 300)-(UserControl.Width - 575, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 90)-(UserControl.Width - 990, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
                picForm.Line (UserControl.Width - 590, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 300), CreateGradientButom(35, 70, lButtonInnerBorder)
        
                ' Button Top Gradient Base Colour
                iVertical = 105
                For I = 1 To 6
                    picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), CreateGradientButom(35, 70, lButtonGradientTop)
                    iVertical = iVertical + 15
                Next
        
                ' Button Bottom Gradient Base Colour
                For I = 1 To 7
                    picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), CreateGradientButom(35, 70, lButtonGradientBottom(I))
                    iVertical = iVertical + 15
                Next
                
                ' Draw the Maximise Button Display
                ' Inside Button Display
                If myForm.Tag <> "vbMaximized" Then
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, &HE9E9E9)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), CreateGradientButom(35, 70, &HE2E2E2)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 705, UserControl.Extender.Top + 225), CreateGradientButom(35, 70, &HDCDCDC)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 705, UserControl.Extender.Top + 240), CreateGradientButom(35, 70, &HD7D7D7)
                    
                    ' Outside Borders
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), CreateGradientButom(35, 70, &H665653)       ' Top Border
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 705, UserControl.Extender.Top + 255), CreateGradientButom(35, 70, &H665653)       ' Bottom Border
                    picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 150)-(UserControl.Width - 870, UserControl.Extender.Top + 255), CreateGradientButom(35, 70, &H665653)       ' Left Border
                    picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 255), CreateGradientButom(35, 70, &H665653)       ' Right Border
            
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 735, UserControl.Extender.Top + 180), CreateGradientButom(35, 70, &H665653)       ' Top Border
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 215)-(UserControl.Width - 735, UserControl.Extender.Top + 215), CreateGradientButom(35, 70, &H665653)       ' Bottom Border
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 825, UserControl.Extender.Top + 215), CreateGradientButom(35, 70, &H665653)       ' Left Border
                    picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 215), CreateGradientButom(35, 70, &H665653)       ' Right Border
                    
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, lButtonGradientBottom(1))
                    ' Outside Borders
                Else
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 150)-(UserControl.Width - 795, UserControl.Extender.Top + 150), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 795, UserControl.Extender.Top + 150)-(UserControl.Width - 720, UserControl.Extender.Top + 150), lButtonGradientBottom(1)
                    picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 150)-(UserControl.Width - 735, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, lButtonGradientBottom(1))
                    
                    picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 120)-(UserControl.Width - 705, UserControl.Extender.Top + 120), CreateGradientButom(35, 70, &H665653)    ' Top Border
                    picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), CreateGradientButom(35, 70, &H665653)    ' Bottom Border
                    picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 120)-(UserControl.Width - 825, UserControl.Extender.Top + 170), CreateGradientButom(35, 70, &H665653)    ' Left Border
                    picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 120)-(UserControl.Width - 705, UserControl.Extender.Top + 225), CreateGradientButom(35, 70, &H665653)    ' Right Border
        
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 180), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 840, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 840, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 840, UserControl.Extender.Top + 210), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 210)-(UserControl.Width - 750, UserControl.Extender.Top + 210), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 840, UserControl.Extender.Top + 225), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 225)-(UserControl.Width - 750, UserControl.Extender.Top + 225), CreateGradientButom(35, 70, &HFFFFFF)
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 750, UserControl.Extender.Top + 240), CreateGradientButom(35, 70, &HFFFFFF)
        
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 750, UserControl.Extender.Top + 165), CreateGradientButom(35, 70, &H665653)    ' Top Border
                    picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 750, UserControl.Extender.Top + 255), CreateGradientButom(35, 70, &H665653)    ' Bottom Border
                    picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 165)-(UserControl.Width - 870, UserControl.Extender.Top + 270), CreateGradientButom(35, 70, &H665653)    ' Left Border
                    picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 165)-(UserControl.Width - 750, UserControl.Extender.Top + 270), CreateGradientButom(35, 70, &H665653)    ' Right Border
                End If
            End If
        End If
   End If
Errhandler:
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picForm_Paint()
    myForm.Picture = Nothing
    ' Paints the Form header and label
    Call SelectColorScheme
    If bPaintForm = False Then
        ' Set the Form header bottom colour
        picForm.ForeColor() = lFormCaptionColor
        Col = lFormGradientBottom
        lBottomR = (Col And &HFF&)
        lBottomG = (Col And &HFF00&) / &H100
        lBottomB = (Col And &HFF0000) / &H10000

        ' Set the Form header top colour
        Col = lFormGradientTop
        lTopR = (Col And &HFF&)
        lTopG = (Col And &HFF00&) / &H100
        lTopB = (Col And &HFF0000) / &H10000

        ' Clear the Form picturebox for drawing and apply the gradient colour
        picForm.Cls
        
        'desativado do modelo origina pois nao funciona em 16 bits de cor
        '''''''Set picForm.Picture = CreateGradient(UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY, False, RGB(lTopR, lTopG, lTopB), RGB(lBottomR, lBottomG, lBottomB), RGBBlend)
        RenderPanel RGB(lTopR, lTopG, lTopB), RGB(lBottomR, lBottomG, lBottomB)
        
        ' Display the Text and Icon
        picForm.FontBold = bFontBold
        picForm.FontItalic = bFontItalic
        picForm.FontSize = dFontSize
        picForm.FontStrikethru = bFontStrikeThru
        picForm.FontUnderline = bFontUnderline
        If bDisplayIcon = False Then
            picForm.CurrentX = 90
        Else
            picForm.ScaleMode = 1
            If imgFormPic.Picture <> 0 Then
                picForm.CurrentX = 330
                If Style_Type = Vista_Aero Then
                    picForm.PaintPicture imgFormPic.Picture, 75, 45, 240, 240
                Else
                    picForm.PaintPicture imgFormPic.Picture, 75, 75, 240, 240
                End If
            Else
                picForm.CurrentX = 90
            End If
        End If
        picForm.CurrentY = (picForm.Height - picForm.TextHeight(sFormCaption)) / 2
        picForm.Print sFormCaption
        
        ' Draw the Buttons
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
        
        ' Top Border Line
        picForm.Line (0, 0)-(picForm.Width - 15, 0), lFormOuterBorder
        
        ' Left Border Line
        picForm.Line (0, 0)-(0, picForm.Height), lFormOuterBorder
        picForm.Line (15, 15)-(15, picForm.Height), lFormInnerBorder
        
        ' Right Border Line
        picForm.Line (picForm.Width - 30, 15)-(picForm.Width - 30, picForm.Height), lFormMiddleBorder
        picForm.Line (picForm.Width - 15, 0)-(picForm.Width - 15, picForm.Height), lFormOuterBorder
        
        If myForm.Tag = "vbNormal" Then
            ' Draw the Mask Colours
            ' Top Left Border Mask
            picForm.Line (0, 0)-(75, 0), &HFF00FF
            picForm.Line (0, 15)-(45, 15), &HFF00FF
            picForm.Line (0, 30)-(30, 30), &HFF00FF
            picForm.Line (0, 45)-(15, 45), &HFF00FF
            picForm.Line (0, 60)-(15, 60), &HFF00FF

            ' Top Left Border
            picForm.Line (45, 15)-(75, 15), lFormOuterBorder
            picForm.Line (30, 30)-(45, 30), lFormOuterBorder
            picForm.Line (15, 45)-(15, 75), lFormOuterBorder
            
            ' Top Right Border Mask
            picForm.Line (picForm.Width - 75, 0)-(picForm.Width, 0), &HFF00FF
            picForm.Line (picForm.Width - 45, 15)-(picForm.Width, 15), &HFF00FF
            picForm.Line (picForm.Width - 30, 30)-(picForm.Width, 30), &HFF00FF
            picForm.Line (picForm.Width - 15, 45)-(picForm.Width, 45), &HFF00FF
            picForm.Line (picForm.Width - 15, 60)-(picForm.Width, 60), &HFF00FF
    
            ' Top Right Border
            picForm.Line (picForm.Width - 75, 15)-(picForm.Width - 45, 15), lFormOuterBorder
            picForm.Line (picForm.Width - 45, 30)-(picForm.Width - 30, 30), lFormOuterBorder
            picForm.Line (picForm.Width - 30, 45)-(picForm.Width - 15, 75), lFormOuterBorder
    
            UserControl.Picture = picForm.Image
            myForm.Picture = UserControl.Picture
            Call MakeTransparent(myForm, &HFF00FF)
        End If
        
        Call UserControl_Paint
        bPaintForm = True
    End If
End Sub

Private Sub TmrMouseMove_Timer()
Static I As Boolean

If GetActiveWindow() <> 0 And I = False Then
    Janela_Ativa = True
'    Debug.Print "A"
    bPaintForm = False
    picForm_Paint
    I = True
ElseIf GetActiveWindow() = 0 And I = True Then
    Janela_Ativa = False
'    Debug.Print "D"
    bPaintForm = False
    picForm_Paint
    I = False
End If
    
    Dim pt As POINTAPI


    ' See where the cursor is.
    GetCursorPos pt

    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> picForm.hwnd Then
        bCloseButtonHover = False
        bCloseButtonClicked = False
        bMinimiseButtonHover = False
        bMaximiseButtonClicked = False
        bMaximiseButtonHover = False
        bMinimiseButtonClicked = False

        If bMouseOnForm = False Then
            ' Draw the Buttons
            Call DrawMinimiseButton
            Call DrawMaximiseButton
            Call DrawCloseButton
            bMouseOnForm = True
        End If
    Else
        bMouseOnForm = False

        ' Draw the Buttons
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
    End If
End Sub

Private Sub UserControl_Initialize()
    
    ' Initialise the default values
    Janela_Ativa = True
    bCloseButton = True
    bCloseButtonClicked = False
    bCloseButtonHover = False
    bDisplayIcon = False
    bFontBold = False
    bFontItalic = False
    bFontStrikeThru = False
    bFontUnderline = False
    bMaximiseButton = True
    bMaximiseButtonClicked = False
    bMaximiseButtonHover = False
    bMinimiseButton = True
    bMinimiseButtonClicked = False
    bMinimiseButtonHover = False
    bMouseClicked = False
    bMouseOnForm = False
    bPaintForm = False
    bRightClick = False
    bSystemTray = False
    bTransparency = False
    bUnloadForm = False
    dFontSize = 8
    iNumControls = 0
    iTransparency = 15
    lFormMaxHeight = 0
    lFormMinHeight = 0
    lFormMaxWidth = 0
    lFormMinWidth = 0
    lSysTrayMenu = 0
    
    ' TwipX and TwipY used only for easier writing
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    ' Following variables used to speed-up the process (prevent recalc of BorderPixels * ...)
    BorderWidth = BorderPixels * TwipX
    BorderHeight = BorderPixels * TwipY
    
End Sub

Private Sub UserControl_InitProperties()
    Call UserControlsCreate
    Set myForm = UserControl.Parent
    myForm.BorderStyle = 0
    myForm.AutoRedraw = True
    
    Call SelectColorScheme
    Call UserControl_Paint
    
    Set myForm.Icon = Nothing
    picForm.Width = picForm.Width - 1
    sFormCaption = UserControl.Extender.Name
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Call moveForm_MouseMove(Button, Shift, X, Y)
    
    Dim Message As Long
    Message = X / Screen.TwipsPerPixelX
    
    Select Case Message
        Case WM_LBUTTONDBLCLK
            moveForm.Visible = Not moveForm.Visible
            moveForm.Tag = Abs(Not moveForm.Visible)
        Case WM_RBUTTONUP
            Dim pt As POINTAPI
            Dim StartX As Single
            Dim StartY As Single
            bRightClick = True
            
            With pt  ' The x & y position where you want the menu displayed
                .X = X
                .Y = Y
                
                GetCursorPos pt
                RaiseEvent Execute(TrackPopupMenuEx(lSysTrayMenu, TPM_RETURNCMD, .X, .Y, myForm.hwnd, ByVal 0&))
            End With
            bRightClick = False
            
            On Error GoTo Errhandler
            If bUnloadForm = True Then
                Call moveForm_Unload(0)
            End If
    End Select
    Exit Sub

Errhandler:
    
End Sub

Private Sub UserControl_Paint()
    On Error GoTo Errhandler
    Set myForm = UserControl.Parent
    myForm.BorderStyle = 0
    
    If bTransparency = True Then Call MakeSemiTransparent(UserControl.Parent.hwnd, iTransparency)
    
    ' Draw the Form Border Lines
    myForm.Cls
    ' Left Form Border Line
    myForm.Line (0, 0)-(0, myForm.Height), lFormOuterBorder
    myForm.Line (15, 0)-(15, myForm.Height), lFormInnerBorder
    myForm.Line (30, 0)-(30, myForm.Height), lFormGradientBottom
    myForm.Line (45, 0)-(45, myForm.Height), lFormGradientBottom
    
    ' Right Border Line
    myForm.Line (myForm.Width - 15, 0)-(myForm.Width - 15, myForm.Height), lFormOuterBorder
    myForm.Line (myForm.Width - 30, 0)-(myForm.Width - 30, myForm.Height), lFormMiddleBorder
    myForm.Line (myForm.Width - 45, 0)-(myForm.Width - 45, myForm.Height), lFormGradientBottom
    myForm.Line (myForm.Width - 60, 0)-(myForm.Width - 60, myForm.Height), lFormGradientBottom
    
    ' Bottom Border Line
    myForm.Line (30, myForm.Height - 60)-(myForm.Width - 30, myForm.Height - 60), lFormGradientBottom
    myForm.Line (30, myForm.Height - 45)-(myForm.Width - 30, myForm.Height - 45), lFormGradientBottom
    myForm.Line (15, myForm.Height - 30)-(myForm.Width - 15, myForm.Height - 30), lFormMiddleBorder
    myForm.Line (0, myForm.Height - 15)-(myForm.Width - 15, myForm.Height - 15), lFormOuterBorder
    
    Dim ScaleSize As Long
    Dim Width, Height As Long 'Width and height of the image on our form
    ScaleSize = myForm.ScaleMode
    
    myForm.ScaleMode = 3
    Width = myForm.ScaleX(myForm.Width, ScaleSize, vbPixels)
    Height = myForm.ScaleY(myForm.Height, ScaleSize, vbPixels)
    
    LoadDragDot Width - 15 + ScaleSize, Height - 15 + ScaleSize
    LoadDragDot Width - 20 + ScaleSize, Height - 10 + ScaleSize
    LoadDragDot Width - 15 + ScaleSize, Height - 10 + ScaleSize
    LoadDragDot Width - 10 + ScaleSize, Height - 10 + ScaleSize
    LoadDragDot Width - 10 + ScaleSize, Height - 15 + ScaleSize
    LoadDragDot Width - 10 + ScaleSize, Height - 20 + ScaleSize
    myForm.ScaleMode = ScaleSize
Errhandler:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moveForm = UserControl.Parent
    Set myForm = UserControl.Parent
    myForm.AutoRedraw = True
    myForm.Tag = "vbNormal"
    FORMRECT.Top = UserControl.Parent.Top
    FORMRECT.Left = UserControl.Parent.Left
    FORMRECT.Width = UserControl.Parent.Width
    FORMRECT.Height = UserControl.Parent.Height
    
    ' Load saved properties
    xlFormGradientTop = PropBag.ReadProperty("ColorFormGradientTop", &H0&)
    xlFormGradientBottom = PropBag.ReadProperty("ColorFormGradientBottom", &H0&)
    
    xlFormInnerBorder = PropBag.ReadProperty("ColorFormInnerBorder", &H0&)
    xlFormMiddleBorder = PropBag.ReadProperty("ColorFormMiddleBorder", &H0&)
    xlFormOuterBorder = PropBag.ReadProperty("ColorFormOuterBorder", &H0&)
        
    xlButtonGradientBottom = PropBag.ReadProperty("ColorButtonGradientBottom", &H0&)
    xlButtonGradientBottomClicked = PropBag.ReadProperty("ColorButtonGradientBottomClicked", &H0&)
    xlButtonGradientBottomHover = PropBag.ReadProperty("ColorButtonGradientBottomHover", &H0&)
    
    xVisualStyles = PropBag.ReadProperty("Style", 0)
    gCTitleDir = PropBag.ReadProperty("StyleBar", 0)
    
    xVisual_Type = PropBag.ReadProperty("Style_Type", 0)
    
    bEnableCloseButton = PropBag.ReadProperty("EnableCloseButton", True)
    bEnableMinimiseButton = PropBag.ReadProperty("EnableMinimiseButton", True)
    bEnableMaximiseButton = PropBag.ReadProperty("EnableMaximiseButton", True)
    
    Call SelectColorScheme
    Call UserControlsCreate
    
    sFormCaption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
    bDisplayIcon = PropBag.ReadProperty("DisplayIcon", False)
    picForm.Font = PropBag.ReadProperty("Font", Ambient.Font)
    bFontBold = PropBag.ReadProperty("FontBold", Ambient.Font)
    bFontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font)
    dFontSize = PropBag.ReadProperty("FontSize", Ambient.Font)
    bFontStrikeThru = PropBag.ReadProperty("FontStrikeThru", Ambient.Font)
    bFontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font)
    lFormCaptionColor = PropBag.ReadProperty("ForeColor", &H0&)
    imgFormPic.Picture = PropBag.ReadProperty("Icon", Nothing)
    lFormMaxHeight = PropBag.ReadProperty("MaxHeight", 0)
    lFormMaxWidth = PropBag.ReadProperty("MaxWidth", 0)
    lFormMinHeight = PropBag.ReadProperty("MinHeight", 0)
    lFormMinWidth = PropBag.ReadProperty("MinWidth", 0)
    bCloseButton = PropBag.ReadProperty("ShowCloseButton", True)
    bMinimiseButton = PropBag.ReadProperty("ShowMinimiseButton", True)
    bMaximiseButton = PropBag.ReadProperty("ShowMaximiseButton", True)
    bSystemTray = PropBag.ReadProperty("ShowSytemTrayIcon", False)
    bTransparency = PropBag.ReadProperty("Transparency", False)
    iTransparency = PropBag.ReadProperty("TransparencyLevel", 15)
    
    Call ShowInTheTaskbar(myForm.hwnd, True)
    If bTransparency = True Then Call MakeSemiTransparent(UserControl.Parent.hwnd, iTransparency)
    picForm.ForeColor = lFormCaptionColor
    
    Call UserControl_Paint
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Call RemoveTransparent(myForm)
    
    If bUnloadForm = False Then
        ' Create the required controls
        Call UserControlsCreate
        
        If Style_Type = Vista_Aero Then
            UserControl.Height = 315
        Else
            UserControl.Height = 390
        End If
         
        UserControl.Extender.Align = 1
       
        ' Position the Form header
        picForm.Move 0, 0, UserControl.Width, UserControl.Height
        picForm.ZOrder 0
        
        bPaintForm = False
        Call picForm_Paint
        bPaintForm = True
    End If
End Sub

Private Sub UserControl_Show()
    myForm.Refresh
End Sub

Private Sub UserControl_Terminate()
    If bUnloadForm = False Then
        myForm.Cls
        Set myForm.Picture = Nothing
    Else
        Unload moveForm
        Unload myForm
        If bSystemTray = True Then
            Call DeleteIconFromTray
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Save properties
    
    Call PropBag.WriteProperty("Caption", sFormCaption, UserControl.Extender.Name)
    Call PropBag.WriteProperty("DisplayIcon", bDisplayIcon, False)
    Call PropBag.WriteProperty("Font", picForm.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", picForm.FontBold, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lFormCaptionColor, &H0&)
    
    Call PropBag.WriteProperty("ColorFormGradientTop", xlFormGradientTop, &H0&)
    Call PropBag.WriteProperty("ColorFormGradientBottom", xlFormGradientBottom, &H0&)
    
    Call PropBag.WriteProperty("ColorFormInnerBorder", xlFormInnerBorder, &H0&)
    Call PropBag.WriteProperty("ColorFormMiddleBorder", xlFormMiddleBorder, &H0&)
    Call PropBag.WriteProperty("ColorFormOuterBorder", xlFormOuterBorder, &H0&)
    
    Call PropBag.WriteProperty("ColorButtonGradientBottom", xlButtonGradientBottom, &H0&)
    Call PropBag.WriteProperty("ColorButtonGradientBottomClicked", xlButtonGradientBottomClicked, &H0&)
    Call PropBag.WriteProperty("ColorButtonGradientBottomHover", xlButtonGradientBottomHover, &H0&)
    
    Call PropBag.WriteProperty("Style_Type", xVisual_Type, 0)
    
    Call PropBag.WriteProperty("EnableCloseButton", bEnableCloseButton, True)
    Call PropBag.WriteProperty("EnableMinimiseButton", bEnableMinimiseButton, True)
    Call PropBag.WriteProperty("EnableMaximiseButton", bEnableMaximiseButton, True)
    
    Call PropBag.WriteProperty("FontItalic", picForm.FontItalic, Ambient.Font)
    Call PropBag.WriteProperty("FontSize", picForm.FontSize, Ambient.Font)
    Call PropBag.WriteProperty("FontStrikethru", picForm.FontStrikethru, Ambient.Font)
    Call PropBag.WriteProperty("FontUnderline", picForm.FontUnderline, Ambient.Font)
    Call PropBag.WriteProperty("Icon", imgFormPic.Picture, Nothing)
    Call PropBag.WriteProperty("MaxHeight", lFormMaxHeight, 0)
    Call PropBag.WriteProperty("MaxWidth", lFormMaxWidth, 0)
    Call PropBag.WriteProperty("MinHeight", lFormMinHeight, 0)
    Call PropBag.WriteProperty("MinWidth", lFormMinWidth, 0)
    Call PropBag.WriteProperty("ShowCloseButton", bCloseButton, True)
    Call PropBag.WriteProperty("ShowMinimiseButton", bMinimiseButton, True)
    Call PropBag.WriteProperty("ShowMaximiseButton", bMaximiseButton, True)
    Call PropBag.WriteProperty("ShowSytemTrayIcon", bSystemTray, False)
    Call PropBag.WriteProperty("StyleBar", gCTitleDir, 0)
    Call PropBag.WriteProperty("Style", xVisualStyles, 0)
    Call PropBag.WriteProperty("Transparency", bTransparency, False)
    Call PropBag.WriteProperty("TransparencyLevel", iTransparency, 15)
End Sub

Public Sub AddSysTrayItem(ByVal MenuID As Long, ByVal MenuCaption As String, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional eItemState As menuEStates)
    If lSysTrayMenu = 0 Then
        lSysTrayMenu = CreatePopupMenu()
    End If
    
    ' Add a SysTray menu item
    If MenuCaption <> "-" Then
        Call AppendMenu(lSysTrayMenu, MFS_STRING Or -bChecked * MFS_CHECKED, MenuID, ByVal MenuCaption)
    Else
        Call AppendMenu(lSysTrayMenu, MFS_STRING Or MFS_SEPARATOR, MenuID, ByVal vbNullString)
    End If
    
    ' Default item
    If bDefault Then Call SetMenuDefaultItem(lSysTrayMenu, MenuID, 0)
    
    ' Disabled (Regular color text)
    If eItemState = xDisabled Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_DISABLED)
    ' Disabled (Disabled color text)
    If eItemState = xGrayed Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_GRAYED)
End Sub

Public Sub AmendSysTrayItem(ByVal MenuID As Long, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional eItemState As menuEStates)
    ' Amend a SysTray menu item
    Dim mnuItemInfo As MENUITEMINFO
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim BuffStr As String * 80   ' Define as largest possible menu text.
    
    With mnuItemInfo
        .cbSize = Len(mnuItemInfo)   ' 44
        .dwTypeData = BuffStr & Chr(0)
        .fType = MFS_STRING
        .cch = Len(mnuItemInfo.dwTypeData)   ' 80
        .fMask = MIIM_STATE
        .wID = MenuID
    End With
    
    ' Get the MenuID original details
    Call GetMenuItemInfo(lSysTrayMenu, MenuID, False, mnuItemInfo)
    
    ' Check or Uncheck
    If bChecked = False Then
        mnuItemInfo.fState = mnuItemInfo.fState - MFS_CHECKED
    Else
        mnuItemInfo.fState = (mnuItemInfo.fState Or MFS_CHECKED)
    End If
    
    ' Default or Not Default
    If bDefault = False Then
        mnuItemInfo.fState = mnuItemInfo.fState - MFS_DEFAULT
    Else
        mnuItemInfo.fState = (mnuItemInfo.fState Or MFS_DEFAULT)
    End If
    
    ' Set the MenuID new details
    Call SetMenuItemInfo(lSysTrayMenu, MenuID, False, mnuItemInfo)
    
    ' Disabled (Regular color text)
    If eItemState = xDisabled Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_DISABLED)
    ' Disabled (Disabled color text)
    If eItemState = xGrayed Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_GRAYED)
    ' Enabled if none of the above
    If eItemState <> xDisabled And eItemState <> xGrayed Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_ENABLED)
End Sub


''1######################################################
Public Property Get ColorFormGradientBottom() As OLE_COLOR
    ColorFormGradientBottom = xlFormGradientBottom
End Property

Public Property Let ColorFormGradientBottom(ByVal New_ForeColor1 As OLE_COLOR)
    xlFormGradientBottom = New_ForeColor1
    PropertyChanged "ColorFormGradientBottom"
    bPaintForm = False
    Call picForm_Paint
End Property


''2######################################################
Public Property Get ColorFormGradientTop() As OLE_COLOR
    ColorFormGradientTop = xlFormGradientTop
End Property

Public Property Let ColorFormGradientTop(ByVal New_ForeColor2 As OLE_COLOR)
    xlFormGradientTop = New_ForeColor2
    PropertyChanged "ColorFormGradientTop"
    bPaintForm = False
    Call picForm_Paint
End Property


''3######################################################
Public Property Get ColorFormOuterBorder() As OLE_COLOR
    ColorFormOuterBorder = xlFormOuterBorder
End Property

Public Property Let ColorFormOuterBorder(ByVal New_ForeColor3 As OLE_COLOR)
    xlFormOuterBorder = New_ForeColor3
    PropertyChanged "ColorFormOuterBorder"
    bPaintForm = False
    Call picForm_Paint
End Property


''4######################################################
Public Property Get ColorFormInnerBorder() As OLE_COLOR
    ColorFormInnerBorder = xlFormInnerBorder
End Property

Public Property Let ColorFormInnerBorder(ByVal New_ForeColor4 As OLE_COLOR)
    xlFormInnerBorder = New_ForeColor4
    PropertyChanged "ColorFormInnerBorder"
    bPaintForm = False
    Call picForm_Paint
End Property


''5######################################################
Public Property Get ColorFormMiddleBorder() As OLE_COLOR
    ColorFormMiddleBorder = xlFormMiddleBorder
End Property

Public Property Let ColorFormMiddleBorder(ByVal New_ForeColor5 As OLE_COLOR)
    xlFormMiddleBorder = New_ForeColor5
    PropertyChanged "ColorFormMiddleBorder"
    bPaintForm = False
    Call picForm_Paint
End Property


''6######################################################
Public Property Get ColorButtonGradientBottom() As OLE_COLOR
    ColorButtonGradientBottom = xlButtonGradientBottom
End Property

Public Property Let ColorButtonGradientBottom(ByVal New_ForeColor6 As OLE_COLOR)
    xlButtonGradientBottom = New_ForeColor6
    
    PropertyChanged "ColorButtonGradientBottom"
    bPaintForm = False
    Call picForm_Paint
End Property


''7######################################################
Public Property Get ColorButtonGradientBottomClicked() As OLE_COLOR
    ColorButtonGradientBottomClicked = xlButtonGradientBottomClicked
End Property

Public Property Let ColorButtonGradientBottomClicked(ByVal New_ForeColor7 As OLE_COLOR)
    xlButtonGradientBottomClicked = New_ForeColor7
    
    PropertyChanged "ColorButtonGradientBottomClicked"
    bPaintForm = False
    Call picForm_Paint
End Property


''8######################################################
Public Property Get ColorButtonGradientBottomHover() As OLE_COLOR
    ColorButtonGradientBottomHover = xlButtonGradientBottomHover
End Property

Public Property Let ColorButtonGradientBottomHover(ByVal New_ForeColor8 As OLE_COLOR)
    xlButtonGradientBottomHover = New_ForeColor8
    
    PropertyChanged "ColorButtonGradientBottomHover"
    bPaintForm = False
    Call picForm_Paint
End Property


''9######################################################
Public Property Get StyleBar() As GRADIENT_DIR1
    StyleBar = gCTitleDir
End Property

Public Property Let StyleBar(Estilo_Degrade As GRADIENT_DIR1)
    gCTitleDir = Estilo_Degrade
    PropertyChanged "StyleBar"

    Call UserControl_Paint
    
    bPaintForm = False
    Call picForm_Paint
End Property


''10######################################################
Public Property Get Style_Type() As xVista_Type
    Style_Type = xVisual_Type
End Property

Public Property Let Style_Type(val1 As xVista_Type)
    xVisual_Type = val1
    PropertyChanged "Style_Type"
    
    Call UserControl_Resize
    Call SelectColorScheme
    Call UserControl_Paint

    bPaintForm = False
    Call picForm_Paint
End Property


''11######################################################
Public Property Get EnableCloseButton() As Boolean
    EnableCloseButton = bEnableCloseButton
End Property

Public Property Let EnableCloseButton(ByVal New_EnableCloseButton As Boolean)
    bEnableCloseButton = New_EnableCloseButton
    PropertyChanged "EnableCloseButton"
    Call UserControl_Resize
End Property


''12######################################################
Public Property Get EnableMinimiseButton() As Boolean
    EnableMinimiseButton = bEnableMinimiseButton
End Property

Public Property Let EnableMinimiseButton(ByVal New_EnableMinimiseButton As Boolean)
    bEnableMinimiseButton = New_EnableMinimiseButton
    PropertyChanged "EnableMinimiseButton"
    Call UserControl_Resize
End Property


''13######################################################
Public Property Get EnableMaximiseButton() As Boolean
    EnableMaximiseButton = bEnableMaximiseButton
End Property

Public Property Let EnableMaximiseButton(ByVal New_EnableMaximiseButton As Boolean)
    bEnableMaximiseButton = New_EnableMaximiseButton
    PropertyChanged "EnableMaximiseButton"
    Call UserControl_Resize
End Property

Public Function FNC_CORES(CORES() As String)
Dim Cor(1 To 32) As String
Cor(1) = "lFormCaptionColor=" & CStr(lFormCaptionColor)
Cor(2) = "lFormGradientBottom=" & CStr(xlFormGradientBottom)
Cor(3) = "lFormGradientTop=" & CStr(xlFormGradientTop)
Cor(4) = "lFormInnerBorder=" & CStr(xlFormInnerBorder)
Cor(5) = "lFormMiddleBorder=" & CStr(xlFormMiddleBorder)
Cor(6) = "lFormOuterBorder=" & CStr(xlFormOuterBorder)
Cor(7) = "lButtonGradientBottom(1)=" & CStr(lButtonGradientBottom(1))
Cor(8) = "lButtonGradientBottom(2)=" & CStr(lButtonGradientBottom(2))
Cor(9) = "lButtonGradientBottom(3)=" & CStr(xlButtonGradientBottom)
Cor(10) = "lButtonGradientBottom(4)=" & CStr(lButtonGradientBottom(4))
Cor(11) = "lButtonGradientBottom(5)=" & CStr(lButtonGradientBottom(5))
Cor(12) = "lButtonGradientBottom(6)=" & CStr(lButtonGradientBottom(6))
Cor(13) = "lButtonGradientBottom(7)=" & CStr(lButtonGradientBottom(7))
Cor(14) = "lButtonGradientBottomClicked(1)=" & CStr(xlButtonGradientBottomClicked)
Cor(15) = "lButtonGradientBottomClicked(2)=" & CStr(lButtonGradientBottomClicked(2))
Cor(16) = "lButtonGradientBottomClicked(3)=" & CStr(lButtonGradientBottomClicked(3))
Cor(17) = "lButtonGradientBottomClicked(4)=" & CStr(lButtonGradientBottomClicked(4))
Cor(18) = "lButtonGradientBottomClicked(5)=" & CStr(lButtonGradientBottomClicked(5))
Cor(19) = "lButtonGradientBottomClicked(6)=" & CStr(lButtonGradientBottomClicked(6))
Cor(20) = "lButtonGradientBottomClicked(7)=" & CStr(lButtonGradientBottomClicked(7))
Cor(21) = "lButtonGradientBottomHover(1)=" & CStr(xlButtonGradientBottomHover)
Cor(22) = "lButtonGradientBottomHover(2)=" & CStr(lButtonGradientBottomHover(2))
Cor(23) = "lButtonGradientBottomHover(3)=" & CStr(lButtonGradientBottomHover(3))
Cor(24) = "lButtonGradientBottomHover(4)=" & CStr(lButtonGradientBottomHover(4))
Cor(25) = "lButtonGradientBottomHover(5)=" & CStr(lButtonGradientBottomHover(5))
Cor(26) = "lButtonGradientBottomHover(6)=" & CStr(lButtonGradientBottomHover(6))
Cor(27) = "lButtonGradientBottomHover(7)=" & CStr(lButtonGradientBottomHover(7))
Cor(28) = "lButtonGradientTop=" & CStr(lFormInnerBorder)
Cor(29) = "lButtonGradientTopClicked=" & CStr(lFormInnerBorder)
Cor(30) = "lButtonGradientTopHover=" & CStr(lFormInnerBorder)
Cor(31) = "lButtonInnerBorder=" & CStr(lFormInnerBorder)
Cor(32) = "lButtonOuterBorder=" & CStr(lFormInnerBorder)

CORES() = Cor
End Function

Private Function CreateGradientButom(Degrau As Byte, Tamanho As Byte, Cor As Long) As Long
Dim RS As Byte, GS As Byte, BS As Byte 'Start RGB
Dim RE As Byte, GE As Byte, BE As Byte 'End RGB
Dim Rc As Byte, GC As Byte, BC As Byte 'Current RGB

RgbCol Cor, RS, GS, BS

Rc = (1& * RS - 255) * ((Tamanho - 1 - Degrau) / (Tamanho - 1)) + 255
GC = (1& * GS - 255) * ((Tamanho - 1 - Degrau) / (Tamanho - 1)) + 255
BC = (1& * BS - 255) * ((Tamanho - 1 - Degrau) / (Tamanho - 1)) + 255

CreateGradientButom = RGB(Rc, GC, BC)
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2@@@@@@@@@@@@@@@@@@@@
'Funçoes retiradas do exemplo abaixo, com essa funcao o barra de titulo
'funciona em 16 bits de cor e 32 bits de cor
'
'Title: Dm PanelFx
'Description: A Stylish Panel controls, supports, Soild Color,
'Texture, Graident, Round Cornners, Moveable, Moveable Partent,
'Titlebar Icon, And a host of other properties.
'This file came from Planet-Source-Code.com...the home millions of '
'lines of source code
'You can view comments on this code/and or vote on it at:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64645&lngWId=1
'The author may have retained certain copyrights to this code...
'please observe their request and the law by reviewing all copyright
'conditions at the above URL.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2@@@@@@@@@@@@@@@@@@@@

Private Sub RenderPanel(StartColor As OLE_COLOR, EndColor As OLE_COLOR)
Dim Rc As RECT
On Error Resume Next
    With picForm
        .Cls
        .BackColor = 0
            SetRect Rc, 0, 0, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY
            GDI_GradientFill .hdc, Rc, StartColor, EndColor, gCTitleDir
            GDI_GradientFill myForm.hdc, Rc, StartColor, EndColor, gCTitleDir
    End With
End Sub

Private Sub GDI_GradientFill(hdc As Long, mRect As RECT, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR1)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX
    'Function used to paint a Gradient effect on the listbox
    
    setTriVertexColor tTV(1), GDI_TranslateColor(mEndColor)
    tTV(0).X = mRect.Left
    tTV(0).Y = mRect.Top
    
    setTriVertexColor tTV(0), GDI_TranslateColor(mStartColor)
    tTV(1).X = mRect.Right
    tTV(1).Y = mRect.Bottom
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
End Sub

Private Function GDI_TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, GDI_TranslateColor) Then
        GDI_TranslateColor = &HFFFF&
    End If
End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, oColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = (oColor And &HFF&) * &H100&
    lGreen = (oColor And &HFF00&)
    lBlue = (oColor And &HFF0000) \ &H100&
    
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef oColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
        oColor = (lComponent And &H7F00&)
        oColor = oColor Or &H8000
    Else
        oColor = lComponent
    End If
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2@@@@@@@@@@@@@@@@@@@@
'Funçoes retiradas do exemplo abaixo, com essa funcao o barra de titulo
'funciona em 16 bits de cor e 32 bits de cor
'
'Title: advanced form skin + Transparent w/o  resource graphics.
'Description: This control does not need any user input. All the user
'has to do is drag and drop it onto a form and your done. This control
'does NOT contain any graphics at all, all pictures are drawn from lines
'and PSET. Many lines of code but all very easy to understand. This control
'allows you to resize, minumize, maximize/restore, and close the form. It
'also reads the forms icon and caption and displays it. Also the HotPink
'areas do become transparent to make this form a little more appealing.
'I plan on addins some animation and better buttons to this project.
'please help me make this better.
'This file came from Planet-Source-Code.com...the home millions of lines ofsource code
'You can view comments on this code/and or vote on it at:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=55882&lngWId=1
'The author may have retained certain copyrights to this code...
'please observe their request and the law by reviewing all copyright conditions at the above URL.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2@@@@@@@@@@@@@@@@@@@@

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Function LoadDragDot(X As Integer, Y As Integer)
    myForm.PSet (X + 0, Y + 0), 16383997
    myForm.PSet (X + 0, Y + 1), 12634052
    myForm.PSet (X + 1, Y + 0), 13025728
    myForm.PSet (X + 1, Y + 1), 393472
    myForm.PSet (X + 1, Y + 2), 16117999
    myForm.PSet (X + 2, Y + 1), 16777214
    myForm.PSet (X + 2, Y + 2), 16776186
End Function

