Attribute VB_Name = "mWndProc"
Option Explicit

Public Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)

' ========================================================

' Code was written in and formatted for 8pt MS San Serif

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
End Enum

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

Public Const NM_FIRST = -0&                ' (0U-  0U)       '  generic to all controls
Public Const NM_CUSTOMDRAW = (NM_FIRST - 12)

Public Enum CD_ReturnFlags
  
' -------------------------------------------------------------------------------------
' When dwDrawStage equals CDDS_PREPAINT:
  
  ' The control will draw itself. It will not send any additional NM_CUSTOMDRAW
  ' messages for this paint cycle.
  CDRF_DODEFAULT = &H0
  ' The control will notify the parent after painting an item.
  CDRF_NOTIFYPOSTPAINT = &H10
  ' The control will notify the parent of any item-related drawing operations. It will send
  ' NM_CUSTOMDRAW notification messages before and after drawing items.
  CDRF_NOTIFYITEMDRAW = &H20
  ' The control will notify the parent after erasing an item.
  CDRF_NOTIFYPOSTERASE = &H40
  ' The control will notify the parent when an item will be erased. It will send
  ' NM_CUSTOMDRAW notification messages before and after erasing items.
  ' no longer supported???!!!
  CDRF_NOTIFYITEMERASE = &H80

' -------------------------------------------------------------------------------------
' When dwDrawStage equals CDDS_ITEMPREPAINT:
  
  ' Your application specified a new font for the item; the control will use the new font.
  CDRF_NEWFONT = &H2

  ' Your application drew the item manually. The control will not draw the item.
  CDRF_SKIPDEFAULT = &H4

  CDRF_NOTIFYSUBITEMDRAW = &H20

End Enum   ' CD_ReturnFlags

' ==================================================================
' NMCUSTOMDRAW structure
Public Type NMCUSTOMDRAW   ' nmcd
  ' An NMHDR structure that contains information about this notification message.
  hdr As NMHDR
  
  ' Specifies the current drawing stage. This value is one of the values below:
  dwDrawStage As CD_DrawStage
  
  ' The handle to the control's device context. Use this HDC to perform any GDI functions.
  hdc As Long
  
  ' A RECT structure that describes the bounding rectangle of the area being drawn.
  rc As RECT
  
  ' The item number. This value is control specific, using the item-referencing
  ' convention for that control. Additionally, trackbar controls use the values below
  ' to identify portions of control.
  dwItemSpec As Long
  
  ' Specifies the current item state. This value is a combination of the flags below.
  uItemState As CD_ItemState
  
  ' Application-defined item data.
  lItemlParam As Long
     
End Type
    
' -------------------------------------------------------------------------------------
'  NMCUSTOMDRAW.dwDrawStage flags:
Public Enum CD_DrawStage

  '  Values under &H10000 are reserved for Global Drawstage Values.
  
  ' Before the painting cycle begins
  CDDS_PREPAINT = &H1

  ' After the painting cycle is complete
  CDDS_POSTPAINT = &H2

  ' Before the erasing cycle begins
  CDDS_PREERASE = &H3

  ' After the erasing cycle is complete
  CDDS_POSTERASE = &H4

  ' The &H10000 bit means it's an Item-Specific Drawstage Value
  
  ' Indicates that the dwItemSpec, uItemState, and lItemParam members are valid.
  CDDS_ITEM = &H10000

  ' Before an item is drawn
  CDDS_ITEMPREPAINT = (CDDS_ITEM Or CDDS_PREPAINT)

  ' After an item has been drawn
  CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)

  ' Before an item is erased
  CDDS_ITEMPREERASE = (CDDS_ITEM Or CDDS_PREERASE)

  ' After an item has been erased
  CDDS_ITEMPOSTERASE = (CDDS_ITEM Or CDDS_POSTERASE)

  CDDS_SUBITEM = &H20000

End Enum   ' CD_DrawStage

' -------------------------------------------------------------------------------------
' NMCUSTOMDRAW.itemState flags:

Public Enum CD_ItemState
  CDIS_SELECTED = &H1    ' The item is selected.
  CDIS_GRAYED = &H2        ' The item is grayed.
  CDIS_DISABLED = &H4     ' The item is disabled.
  CDIS_CHECKED = &H8      ' The item is checked.
  CDIS_FOCUS = &H10         ' The item is in focus.
  CDIS_DEFAULT = &H20     ' The item is in its default state.
  CDIS_HOT = &H40              ' The item is currently under the pointer ("hot").
  CDIS_MARKED = &H80      ' The item is marked. The meaning of this is up to the implementation.
  CDIS_INDETERMINATE = &H100   ' The item is in an indeterminate state.
End Enum

' ==================================================================
' Listview
Public Type NMLVCUSTOMDRAW
  
  ' NMCUSTOMDRAW structure that contains general Custom Draw information.
  nmcd As NMCUSTOMDRAW
  
  ' A COLORREF value representing the color that will be used to display text
  ' foreground in the list view control.
  clrText As Long
  
  ' A COLORREF value representing the color that will be used to display text
  ' background in the list view control.
  clrTextBk As Long

  iSubItem As Long

End Type

Private SpecialLV As ListView

'---------------------------------------------------------------------------------------
' Procedure : SubClass
' DateTime  : 4/10/2004 10:47
' Purpose   : Para leitura de eventos de listview
' Inputs    :
' Outputs   :
'---------------------------------------------------------------------------------------
Public Sub SubClassLV(ByVal hWnd As Long, ByVal lpfnNew As Long, LV As ListView)
Dim Sucesso As Boolean
Dim lpfnOld As Long
  
  On Error GoTo Falha
  
    'acerta lista que vai ser alterada
    Set SpecialLV = LV
  
    'se ja subclassed
    If GetProp(hWnd, OLDWNDPROC) Then
        Exit Sub
    End If
  
    'pega nova tela ou propriedade
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)
    If lpfnOld Then
        Sucesso = SetProp(hWnd, OLDWNDPROC, lpfnOld)
    End If
  
Falha:
    'se erro = volta com antigo
    If Not Sucesso And lpfnOld Then
        SetWindowLong hWnd, GWL_WNDPROC, lpfnOld
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UnSubClass
' DateTime  : 4/10/2004 10:49
' Purpose   : Remove leitura de eventos de form de supervisao
' Inputs    :
' Outputs   :
'---------------------------------------------------------------------------------------
Public Sub UnSubClassLV(ByVal hWnd As Long)
Dim lpfnOld As Long
  
    'remove subclass
    lpfnOld = GetProp(hWnd, OLDWNDPROC)
    If lpfnOld Then
        If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
            RemoveProp hWnd, OLDWNDPROC
            RemoveProp hWnd, OBJECTPTR
        End If
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : WndProcLV
' DateTime  : 4/10/2004 10:50
' Purpose   : faz controle de pintura de tela
' Inputs    :
' Outputs   :
'---------------------------------------------------------------------------------------
Public Function WndProcLV(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Item As MSComctlLib.ListItem
Dim i As Byte

On Error Resume Next

  Select Case uMsg
    
    ' ======================================================
    Case WM_NOTIFY
    Static lvcd As NMLVCUSTOMDRAW
    
      MoveMemory lvcd, ByVal lParam, Len(lvcd)   ' every structs' member is a Long
      Select Case lvcd.nmcd.hdr.code
         
        Case NM_CUSTOMDRAW
          Static iElement As Long
          
          Select Case lvcd.nmcd.dwDrawStage
          
            ' ====================================================
            Case CDDS_PREPAINT
              ' Tell the listview we want CDDS_ITEMPREPAINT for each item
                WndProcLV = CDRF_NOTIFYITEMDRAW
                Exit Function
  
            ' ====================================================
            Case CDDS_ITEMPREPAINT
                Set Item = SpecialLV.ListItems(CInt(lvcd.nmcd.dwItemSpec) + 1)
                i = InStr(1, Item.Tag, " ")
                lvcd.clrText = CLng(Trim$(Mid$(Item.Tag, i)))
                lvcd.clrTextBk = CLng(Trim$(Left$(Item.Tag, i)))
                MoveMemory ByVal lParam, lvcd, Len(lvcd)
                Set Item = Nothing
              
              WndProcLV = CDRF_NOTIFYSUBITEMDRAW Or CDRF_NEWFONT
              Exit Function
  
            Case (CDDS_ITEMPREPAINT Or CDDS_SUBITEM)
                Set Item = SpecialLV.ListItems(CInt(lvcd.nmcd.dwItemSpec) + 1)
                i = InStr(1, Item.Tag, " ")
                lvcd.clrText = CLng(Trim$(Mid$(Item.Tag, i)))
                lvcd.clrTextBk = CLng(Trim$(Left$(Item.Tag, i)))
                MoveMemory ByVal lParam, lvcd, Len(lvcd)
                WndProcLV = CDRF_NEWFONT
              Exit Function
  
          End Select
      End Select
                
    ' ======================================================
    ' Unsubclass the window.
    Case WM_DESTROY
      CallWindowProc GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam
      UnSubClassLV hWnd
      Exit Function
      
  End Select   ' uMsg
  
  WndProcLV = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)

End Function
