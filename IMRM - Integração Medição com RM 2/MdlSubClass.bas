Attribute VB_Name = "MdlSubClass"
Option Explicit
Private Const GWL_WNDPROC = (-4)
Private Enum EErrorWindowProc
   eeBaseWindowProc = 13080
   eeCantSubclass
   eeAlreadyAttached
   eeInvalidWindow
   eeNoExternalWindow
End Enum
Private Declare Function IsWindow Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetProp _
                Lib "user32" _
                Alias "GetPropA" (ByVal Hwnd As Long, _
                                  ByVal lpString As String) As Long
Private Declare Function SetProp _
                Lib "user32" _
                Alias "SetPropA" (ByVal Hwnd As Long, _
                                  ByVal lpString As String, _
                                  ByVal hData As Long) As Long
Private Declare Function RemoveProp _
                Lib "user32" _
                Alias "RemovePropA" (ByVal Hwnd As Long, _
                                     ByVal lpString As String) As Long
Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal Hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal Hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal Hwnd As Long, _
                              lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (lpvDest As Any, _
                                       lpvSource As Any, _
                                       ByVal cbCopy As Long)
Private m_iCurrentMessage As Long
Private m_iProcOld        As Long
Public Property Get CurrentMessage() As Long
   CurrentMessage = m_iCurrentMessage
End Property
Sub AttachMessage(iwp As ISubclass, ByVal Hwnd As Long, ByVal iMsg As Long)
   Dim procOld As Long, F As Long, c As Long
   Dim iC      As Long, bFail As Boolean
   If IsWindow(Hwnd) = False Then ErrRaise eeInvalidWindow
   If IsWindowLocal(Hwnd) = False Then ErrRaise eeNoExternalWindow
   c = GetProp(Hwnd, "C" & Hwnd)
   If c = 0 Then
      procOld = SetWindowLong(Hwnd, GWL_WNDPROC, AddressOf WindowProc)
      If procOld = 0 Then ErrRaise eeCantSubclass
      F = SetProp(Hwnd, Hwnd, procOld)
      Debug.Assert F <> 0
      c = 1
      F = SetProp(Hwnd, "C" & Hwnd, c)
   Else
      c = c + 1
      F = SetProp(Hwnd, "C" & Hwnd, c)
   End If
   Debug.Assert F <> 0
   c = GetProp(Hwnd, Hwnd & "#" & iMsg & "C")
   If (c > 0) Then
      For iC = 1 To c
         If (GetProp(Hwnd, Hwnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
            ErrRaise eeAlreadyAttached
            bFail = True
            Exit For
         End If
      Next iC
   End If
   If Not (bFail) Then
      c = c + 1
      F = SetProp(Hwnd, Hwnd & "#" & iMsg & "C", c)
      Debug.Assert F <> 0
      F = SetProp(Hwnd, Hwnd & "#" & iMsg & "#" & c, ObjPtr(iwp))
      Debug.Assert F <> 0
   End If
End Sub
Public Function CallOldWindowProc(ByVal Hwnd As Long, _
                                  ByVal iMsg As Long, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long) As Long
   CallOldWindowProc = CallWindowProc(m_iProcOld, Hwnd, iMsg, wParam, lParam)
End Function
Sub DetachMessage(iwp As ISubclass, ByVal Hwnd As Long, ByVal iMsg As Long)
   Dim procOld As Long, F As Long, c As Long
   Dim iC      As Long, iP As Long, lPtr As Long
   c = GetProp(Hwnd, "C" & Hwnd)
   If c = 1 Then
      procOld = GetProp(Hwnd, Hwnd)
      Debug.Assert procOld <> 0
      Call SetWindowLong(Hwnd, GWL_WNDPROC, procOld)
      RemoveProp Hwnd, Hwnd
      RemoveProp Hwnd, "C" & Hwnd
   Else
      c = GetProp(Hwnd, "C" & Hwnd)
      c = c - 1
      F = SetProp(Hwnd, "C" & Hwnd, c)
   End If
   c = GetProp(Hwnd, Hwnd & "#" & iMsg & "C")
   If (c > 0) Then
      For iC = 1 To c
         If (GetProp(Hwnd, Hwnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
            iP = iC
            Exit For
         End If
      Next iC
      If (iP <> 0) Then
         For iC = iP + 1 To c
            lPtr = GetProp(Hwnd, Hwnd & "#" & iMsg & "#" & iC)
            SetProp Hwnd, Hwnd & "#" & iMsg & "#" & (iC - 1), lPtr
         Next iC
      End If
      RemoveProp Hwnd, Hwnd & "#" & iMsg & "#" & c
      c = c - 1
      SetProp Hwnd, Hwnd & "#" & iMsg & "C", c
   End If
End Sub
Function IsWindowLocal(ByVal Hwnd As Long) As Boolean
   Dim idWnd As Long
   Call GetWindowThreadProcessId(Hwnd, idWnd)
   IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function
Private Sub ErrRaise(e As Long)
   Dim sText As String, sSource As String
   If e > 1000 Then
      sSource = App.EXEName & ".WindowProc"
      Select Case e
         Case eeCantSubclass
            sText = "Can't subclass window"
         Case eeAlreadyAttached
            sText = "Message already handled by another class"
         Case eeInvalidWindow
            sText = "Invalid window"
         Case eeNoExternalWindow
            sText = "Can't modify external window"
      End Select
      Err.Raise e Or vbObjectError, sSource, sText
   Else
      Err.Raise e, sSource
   End If
End Sub
Private Function WindowProc(ByVal Hwnd As Long, _
                            ByVal iMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
   Dim procOld As Long, pSubclass As Long
   Dim iwp     As ISubclass, iwpT As ISubclass
   Dim iPC     As Long, iP As Long, bNoProcess As Long
   Dim bCalled As Boolean
   procOld = GetProp(Hwnd, Hwnd)
   Debug.Assert procOld <> 0
   bCalled = False
   iPC = GetProp(Hwnd, Hwnd & "#" & iMsg & "C")
   If (iPC > 0) Then
      For iP = 1 To iPC
         bNoProcess = False
         pSubclass = GetProp(Hwnd, Hwnd & "#" & iMsg & "#" & iP)
         If pSubclass = 0 Then
            WindowProc = CallWindowProc(procOld, Hwnd, iMsg, wParam, ByVal lParam)
            bNoProcess = True
         End If
         If Not (bNoProcess) Then
            CopyMemory iwpT, pSubclass, 4
            Set iwp = iwpT
            CopyMemory iwpT, 0&, 4
            m_iCurrentMessage = iMsg
            m_iProcOld = procOld
            With iwp
               If (iP = 1) Then
                  If .MsgResponse = emrPreprocess Then
                     If Not (bCalled) Then
                        WindowProc = CallWindowProc(procOld, Hwnd, iMsg, wParam, ByVal lParam)
                        bCalled = True
                     End If
                  End If
               End If
               WindowProc = .WindowProc(Hwnd, iMsg, wParam, ByVal lParam)
               If (iP = iPC) Then
                  If .MsgResponse = emrPostProcess Then
                     If Not (bCalled) Then
                        WindowProc = CallWindowProc(procOld, Hwnd, iMsg, wParam, ByVal lParam)
                        bCalled = True
                     End If
                  End If
               End If
            End With
         End If
      Next iP
   Else
      WindowProc = CallWindowProc(procOld, Hwnd, iMsg, wParam, ByVal lParam)
   End If
End Function
