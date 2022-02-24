Attribute VB_Name = "Bordas"
Option Explicit

Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_DRAWFRAME Or SWP_NOSIZE Or SWP_NOMOVE

Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Para poder usar este efeito use a função abaixo

Public Sub BordasControle(gForm As Form, ctrl As Control, b As Boolean)
''''''''parametros
''''''''gForm->formulario dos objetos
''''''''ctrl->controles
''''''''b->true-redimensionavel
    Dim lngStyle As Long
    Dim X As Long

    lngStyle = GetWindowLong(ctrl.hwnd, GWL_STYLE)

    If b Then
        lngStyle = lngStyle Or WS_THICKFRAME
    Else
        lngStyle = lngStyle Xor WS_THICKFRAME
    End If

    X = SetWindowLong(ctrl.hwnd, GWL_STYLE, lngStyle)
    X = SetWindowPos(ctrl.hwnd, gForm.hwnd, 0, 0, 0, 0, SWP_FLAGS)
End Sub
