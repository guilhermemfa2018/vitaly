Attribute VB_Name = "ModRezize"
'Created:       07/10/97
'Author:        David Thieme

Private Type ctrObj
    Name As String
    Index As Long
    Parrent As String
    Top As Long
    Left As Long
    Height As Long
    Width As Long
    ScaleHeight As Long
    ScaleWidth As Long
End Type

Private FormRecord() As ctrObj
Private ControlRecord() As ctrObj
Private MaxForm As Long
Private MaxControl As Long

Private Function ActualPos(plLeft As Long) As Long
    If plLeft < 0 Then
        ActualPos = plLeft + 75000
    Else
        ActualPos = plLeft
    End If
End Function

Private Function FindForm(pfrmIn As Form) As Long

Dim i As Long
    FindForm = -1
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                FindForm = i
                Exit Function
            End If
        Next i
    End If
End Function

Private Function AddForm(pfrmIn As Form) As Long
Dim FormControl As Control
Dim i As Long
    ReDim Preserve FormRecord(MaxForm + 1)
    FormRecord(MaxForm).Name = pfrmIn.Name
    FormRecord(MaxForm).Top = pfrmIn.Top
    FormRecord(MaxForm).Left = pfrmIn.Left
    FormRecord(MaxForm).Height = pfrmIn.Height
    FormRecord(MaxForm).Width = pfrmIn.Width
    FormRecord(MaxForm).ScaleHeight = pfrmIn.ScaleHeight
    FormRecord(MaxForm).ScaleWidth = pfrmIn.ScaleWidth
    AddForm = MaxForm
    MaxForm = MaxForm + 1
    For Each FormControl In pfrmIn
        i = FindControl(FormControl, pfrmIn.Name)
        If i < 0 Then
            i = AddControl(FormControl, pfrmIn.Name)
        End If
    Next FormControl
End Function

Private Function FindControl(inControl As Control, inName As String) As Long
Dim i As Long
    FindControl = -1
    For i = 0 To (MaxControl - 1)
        If ControlRecord(i).Parrent = inName Then
            If ControlRecord(i).Name = inControl.Name Then
                On Error Resume Next
                If ControlRecord(i).Index = inControl.Index Then
                    FindControl = i
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Function

Private Function AddControl(inControl As Control, inName As String) As Long
    ReDim Preserve ControlRecord(MaxControl + 1)
    On Error Resume Next
    ControlRecord(MaxControl).Name = inControl.Name
    ControlRecord(MaxControl).Index = inControl.Index
    ControlRecord(MaxControl).Parrent = inName
    If TypeOf inControl Is Line Then
        ControlRecord(MaxControl).Top = inControl.Y1
        ControlRecord(MaxControl).Left = ActualPos(inControl.X1)
        ControlRecord(MaxControl).Height = inControl.Y2
        ControlRecord(MaxControl).Width = ActualPos(inControl.X2)
    Else
        ControlRecord(MaxControl).Top = inControl.Top
        ControlRecord(MaxControl).Left = ActualPos(inControl.Left)
        ControlRecord(MaxControl).Height = inControl.Height
        ControlRecord(MaxControl).Width = inControl.Width
    End If
    'If TypeOf inControl Is DBList Then
    '    inControl.IntegralHeight = False
    'End If
    On Error GoTo 0
    AddControl = MaxControl
    MaxControl = MaxControl + 1
End Function

Private Function PerWidth(pfrmIn As Form) As Long
Dim i As Long
    i = FindForm(pfrmIn)
    If i < 0 Then
        i = AddForm(pfrmIn)
    End If
    PerWidth = (pfrmIn.ScaleWidth * 100) \ FormRecord(i).ScaleWidth
End Function

Private Function PerHeight(pfrmIn As Form) As Single
Dim i As Long
    i = FindForm(pfrmIn)
    If i < 0 Then
        i = AddForm(pfrmIn)
    End If
    PerHeight = (pfrmIn.ScaleHeight * 100) \ FormRecord(i).ScaleHeight
End Function

Private Sub ResizeControl(inControl As Control, pfrmIn As Form)
Dim i As Long
Dim yRatio, xRatio, lTop, lLeft, lWidth, lHeight As Long
    yRatio = PerHeight(pfrmIn)
    xRatio = PerWidth(pfrmIn)
    i = FindControl(inControl, pfrmIn.Name)
    On Error GoTo Moveit
    If inControl.Left < 0 Then
        lLeft = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
    Else
        lLeft = CLng((ControlRecord(i).Left * xRatio) \ 100)
    End If
    lTop = CLng((ControlRecord(i).Top * yRatio) \ 100)
    lWidth = CLng((ControlRecord(i).Width * xRatio) \ 100)
    lHeight = CLng((ControlRecord(i).Height * yRatio) \ 100)
    GoTo Moveit
Moveit:
    On Error GoTo MoveError1
    If TypeOf inControl Is Line Then
        If inControl.X1 < 0 Then
            inControl.X1 = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
        Else
            inControl.X1 = CLng((ControlRecord(i).Left * xRatio) \ 100)
        End If
        inControl.Y1 = CLng((ControlRecord(i).Top * yRatio) \ 100)
        If inControl.X2 < 0 Then
            inControl.X2 = CLng(((ControlRecord(i).Width * xRatio) \ 100) - 75000)
        Else
            inControl.X2 = CLng((ControlRecord(i).Width * xRatio) \ 100)
        End If
        inControl.Y2 = CLng((ControlRecord(i).Height * yRatio) \ 100)
    Else
        If TypeOf inControl Is Timer Then
            GoTo subExit
        End If
        If TypeOf inControl Is Image Then  ' ImageList
            GoTo subExit
        End If
        'If TypeOf inControl Is CrystalReport Then
        '    GoTo subExit
        'End If
        If TypeOf inControl Is Skin Then
            GoTo subExit
        End If
        'If TypeOf inControl Is CommonDialog Then
        '    GoTo subExit
        'End If
        On Error Resume Next
        inControl.Move lLeft, lTop, lWidth, lHeight
    End If
    GoTo subExit
MoveError1:
    On Error GoTo MoveError2
    inControl.Move lLeft, lTop, lWidth
    GoTo subExit
MoveError2:
    On Error GoTo subExit
    inControl.Move lLeft, lTop
subExit:
    On Error GoTo 0
End Sub

Public Sub ResizeForm(gForm As Form)
Dim FormControl As Control
Dim isVisible As Boolean
If gForm.Top < 30000 Then
    isVisible = gForm.Visible
    gForm.Visible = False
    For Each FormControl In gForm
        Call ResizeControl(FormControl, gForm)
    Next FormControl
    gForm.Visible = isVisible
End If
End Sub

Public Sub SaveFormPosition(gForm As Form)
Dim i As Long
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = gForm.Name Then
                FormRecord(i).Top = gForm.Top
                FormRecord(i).Left = gForm.Left
                FormRecord(i).Height = gForm.Height
                FormRecord(i).Width = gForm.Width
                Exit Sub
            End If
        Next i
        AddForm (gForm)
    End If
End Sub

Private Sub RestoreFormPosition(gForm As Form)
Dim i As Long
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = gForm.Name Then
                If FormRecord(i).Top < 0 Then
                    gForm.WindowState = 2
                ElseIf FormRecord(i).Top < 30000 Then
                    gForm.WindowState = 0
                    gForm.Move FormRecord(i).Left, FormRecord(i).Top, FormRecord(i).Width, FormRecord(i).Height
                Else
                    gForm.WindowState = 1
                End If
                Exit Sub
            End If
        Next i
    End If
End Sub


