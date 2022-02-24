Attribute VB_Name = "CorDBGrid"
'"WDS - Autor Weber, 14-01-06 - adptação do RGB"
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type
Private Type GRADIENT_RECT
    UPPERLEFT  As Long
    LOWERRIGHT As Long
End Type
Private Type TRIVERTEX
    X       As Long
    Y       As Long
    Red     As Integer
    Green   As Integer
    Blue    As Integer
    Alpha   As Integer
End Type

Private Type RGB
    R As Integer
    g As Integer
    b As Integer
End Type
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal HWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal HWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal HWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal HWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal HWnd As Long, ByVal lpRect As Long) As Long
Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_PAINT    As Long = &HF
Private Const WM_DESTROY  As Long = &H2
Private Const WM_TIMER    As Long = &H113
Private Const ID_TIMER    As Long = &HCBABE
Public Enum TabStyle
       cSolidColor = 0
       cPicture = 1
       cGradient = 2
       cAnimatedGradient = 3
End Enum
Public Enum Direction
       cHorizontal = 0
       cVertical = 1
End Enum
Private DestDC      As Long
Private maskDC      As Long
Private MemDC       As Long
Private OrigDC      As Long
Private MaskPic     As Long
Private MemPic      As Long
Private TempPic     As Long
Private OrigPic     As Long
Private TempDC      As Long
Private origBrush As Long
Private TempBrush As Long
Private origColor As Long
Private gColor1   As Long
Private gColor2   As Long
Private gDir      As Long
Private gTime     As Long
Private gFadeFlag As Boolean
Private ImageWidth  As Long
Private ImageHeight As Long
Private oldWndProc As Long
Private Function GetLngColor(Color As Long) As Long
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function
Private Function GetRGBColors(Color As Long) As RGB
Dim HexColor As String
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.R = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.g = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.b = "&H" & Mid(HexColor, 1, 2) & "00"
End Function
Public Sub SetStyle(ByVal HWnd As Long, ByRef Style As TabStyle)
           SetProp HWnd, "MyStyle", Style
End Sub
Public Sub SetFadeTime(ByVal HWnd As Long, ByVal cTime As Long)
    If cTime > 10 Then cTime = 10
    If cTime < 1 Then cTime = 1
           SetProp HWnd, "MyFadeTime", cTime
End Sub
Private Function GetFadeTime(ByVal HWnd As Long) As Long
           GetFadeTime = GetProp(HWnd, "MyFadeTime")
End Function
Private Function GetStyleParams(ByVal HWnd As Long) As TabStyle
           GetStyleParams = GetProp(HWnd, "MyStyle")
End Function
Public Sub SetGradientDir(ByVal HWnd As Long, ByRef Style As Direction)
           SetProp HWnd, "MyGradientDir", Style
End Sub
Private Sub GetGradientDir(ByVal HWnd As Long)
           gDir = GetProp(HWnd, "MyGradientDir")
End Sub
Private Sub SetHookInstance(ByVal HWnd As Long)
           SetProp HWnd, "Hooked", True
End Sub
Private Function CheckHookInstance(ByVal HWnd As Long) As Boolean
           CheckHookInstance = GetProp(HWnd, "Hooked")
End Function
Public Sub SetSolidColor(ByVal HWnd As Long, ByVal Color As Long)
           SetProp HWnd, "MySolidColor", GetLngColor(Color)
End Sub
Public Sub SetGradientColor1(ByVal HWnd As Long, ByVal Color As Long)
           SetProp HWnd, "MyGradientColor1", GetLngColor(Color)
End Sub
Public Sub SetGradientColor2(ByVal HWnd As Long, ByVal Color As Long)
           SetProp HWnd, "MyGradientColor2", GetLngColor(Color)
End Sub
Private Sub GetSolidColor(ByVal HWnd As Long)
     TempBrush = CreateSolidBrush(GetProp(HWnd, "MySolidColor"))
End Sub
Private Sub GetGradientColor1(ByVal HWnd As Long)
     gColor1 = GetProp(HWnd, "MyGradientColor1")
End Sub
Private Sub GetGradientColor2(ByVal HWnd As Long)
     gColor2 = GetProp(HWnd, "MyGradientColor2")
End Sub
Public Sub SetPicture(ByVal HWnd As Long, ByVal Width As Long, ByVal Height As Long, ByRef cPicture As StdPicture)
           SetProp HWnd, "MyPicture", cPicture.Handle
           SetProp HWnd, "MyPictureWidth", Width
           SetProp HWnd, "MyPictureHeight", Height
End Sub
Private Sub GetPictureParams(ByVal HWnd As Long)
    TempBrush = CreatePatternBrush(GetProp(HWnd, "MyPicture"))
    ImageWidth = GetProp(HWnd, "MyPictureWidth")
    ImageHeight = GetProp(HWnd, "MyPictureHeight")
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub SelectBitmap()
Dim cHandle As Long
       cHandle = SelectObject(maskDC, MaskPic)
       DeleteObject cHandle
       cHandle = SelectObject(MemDC, MemPic)
       DeleteObject cHandle
       cHandle = SelectObject(TempDC, TempPic)
       DeleteObject cHandle
       cHandle = SelectObject(OrigDC, OrigPic)
       DeleteObject cHandle
End Sub
Private Sub CreateBackMask(ByVal m_Width As Long, ByVal m_Height As Long)
        origColor = SetBkColor(DestDC, GetSysColor(15))
        SetBkColor OrigDC, GetSysColor(15)
        BitBlt maskDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcCopy
End Sub
Private Sub CreateNewDCWorkArea(ByVal m_Width As Long, ByVal m_Height As Long)
        maskDC = CreateCompatibleDC(DestDC)
        MaskPic = CreateBitmap(m_Width, m_Height, 1, 1, ByVal 0&)
        MemDC = CreateCompatibleDC(DestDC)
        MemPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        TempDC = CreateCompatibleDC(DestDC)
        TempPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        OrigDC = CreateCompatibleDC(DestDC)
        OrigPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
End Sub
Private Sub DOBitBlt(ByVal m_Width As Long, ByVal m_Height As Long)
        BitBlt MemDC, 0, 0, m_Width, m_Height, maskDC, 0, 0, vbSrcCopy
        BitBlt MemDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcPaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, maskDC, 0, 0, vbMergePaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MemDC, 0, 0, vbSrcAnd
        BitBlt DestDC, 0, 0, m_Width, m_Height, TempDC, 0, 0, vbSrcCopy
End Sub
Private Sub CleanDCs()
        DeleteDC TempDC
        DeleteObject TempPic
        DeleteDC maskDC
        DeleteObject MaskPic
        DeleteDC MemDC
        DeleteObject MemPic
        DeleteDC OrigDC
        DeleteObject OrigPic
        DeleteObject TempBrush
End Sub
Private Sub DrawGradient(cHdc As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color1 As RGB, Color2 As RGB, Optional Direction = 1)
Dim Vert(1) As TRIVERTEX
Dim gRect As GRADIENT_RECT
    With Vert(0)
        .X = X
        .Y = Y
        .Red = Color1.R
        .Green = Color1.g
        .Blue = Color1.b
        .Alpha = 0&
    End With
    With Vert(1)
        .X = Vert(0).X + X2
        .Y = Vert(0).Y + Y2
        .Red = Color2.R
        .Green = Color2.g
        .Blue = Color2.b
        .Alpha = 0&
    End With
    gRect.UPPERLEFT = 0
    gRect.LOWERRIGHT = 1
    GradientFillRect cHdc, Vert(0), 2, gRect, 1, Direction
End Sub
Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long
Dim R As Long
Dim g As Long
Dim b As Long
      R = (Color And &HFF) + Value
      g = ((Color \ &H100) Mod &H100) + Value
      b = ((Color \ &H10000) Mod &H100)
      b = b + ((b * Value) \ &HC0)
    If Value > 0 Then
        If R > 255 Then R = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
    ElseIf Value < 0 Then
        If R < 0 Then R = 0
        If g < 0 Then g = 0
        If b < 0 Then b = 0
    End If
    ShiftColor = R + 256& * g + 65536 * b
End Function

'######### - SSTab - ############
Public Sub SSTabSubclass(ByVal HWnd As Long)
If Not CheckHookInstance(HWnd) Then
    SetHookInstance HWnd
    oldWndProc = SetWindowLong(HWnd, GWL_WNDPROC, AddressOf oldSSTabProc)
End If
End Sub
Public Function oldSSTabProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
       If GetStyleParams(HWnd) = cAnimatedGradient Then
          KillTimer HWnd, 0
          SetTimer HWnd, ID_TIMER, 1, 0
       End If
       oldSSTabProc = NewSSTabProc(HWnd, uMsg, wParam, lParam)
End Function
Private Function NewSSTabProc(ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     On Error Resume Next
    Dim m_ItemRect As RECT
    Dim m_Width    As Long
    Dim m_Height   As Long
    If wMsg = WM_PAINT Then
        DestDC = GetDC(HWnd)
        GetWindowRect HWnd, m_ItemRect
                m_Width = m_ItemRect.right - m_ItemRect.Left
                m_Height = m_ItemRect.bottom - m_ItemRect.Top
        Select Case GetStyleParams(HWnd)
         Case cPicture
              GetPictureParams HWnd
         Case cSolidColor
              GetSolidColor HWnd
         Case cGradient
              GetGradientColor1 HWnd
              GetGradientColor2 HWnd
              GetGradientDir HWnd
         Case cAnimatedGradient
              GetGradientDir HWnd
         
         Case Else
               Debug.Print "Invalid Style"
        End Select
        CreateNewDCWorkArea m_Width, m_Height
        Call SelectBitmap
        CallWindowProc oldWndProc, HWnd, wMsg, OrigDC, lParam
        Call CreateBackMask(m_Width, m_Height)
        origBrush = SelectObject(TempDC, TempBrush)
        If GetStyleParams(HWnd) = cGradient Or GetStyleParams(HWnd) = cAnimatedGradient Then
            DrawGradient TempDC, 0, 0, m_Width, m_Height, GetRGBColors(gColor1), GetRGBColors(gColor2), gDir
        Else
            PatBlt TempDC, 0, 0, m_Width, m_Height, vbPatCopy
        End If
        SelectObject TempDC, origBrush
        Call DOBitBlt(m_Width, m_Height)
        Call CleanDCs
        SetBkColor DestDC, origColor
        ReleaseDC HWnd, DestDC
        ValidateRect HWnd, 0
    ElseIf wMsg = WM_TIMER Then
        If GetStyleParams(HWnd) <> cAnimatedGradient Then
            KillTimer HWnd, 0
            Exit Function
        End If
        If gFadeFlag Then
            gTime = gTime - GetFadeTime(HWnd)
        Else
            gTime = gTime + GetFadeTime(HWnd)
        End If
        If gTime > 255 Then
           gTime = 255
           gFadeFlag = Not gFadeFlag
        ElseIf gTime < 0 Then
           gTime = 0
           gFadeFlag = Not gFadeFlag
        End If
        GetGradientColor1 HWnd
        GetGradientColor2 HWnd
        gColor1 = ShiftColor(gColor1, gTime)
        gColor2 = ShiftColor(gColor2, gTime)
        RedrawWindow HWnd, ByVal 0&, ByVal 0&, &H1
        Debug.Print gTime
    ElseIf wMsg = WM_DESTROY Then
        KillTimer HWnd, 0
        DeleteObject TempBrush
        SetWindowLong HWnd, GWL_WNDPROC, oldWndProc
        NewSSTabProc = CallWindowProc(oldWndProc, HWnd, wMsg, wParam, lParam)
    Else
        NewSSTabProc = CallWindowProc(oldWndProc, HWnd, wMsg, wParam, lParam)
    End If
End Function

'##################################################################################
Public Sub NewColorDBGrid(Formu As Form)

    Dim controle As Control
   
    On Error Resume Next
    
    For Each controle In Formu.Controls
        If TypeOf controle Is DBGrid Then
            'função para converter a cor do DBGrid
            SetStyle controle.HWnd, cSolidColor
            'este ultimo parametro é q define a cor
            SetSolidColor controle.HWnd, Principal.Skin1.WindowColor
            SSTabSubclass controle.HWnd
            
            controle.BackColor = Principal.Skin1.WindowColor
            controle.Font = "tahoma"
            controle.FontSize = 9
            controle.FontBold = False
            controle.ForeColor = &H0&
            
        End If
    Next
End Sub
