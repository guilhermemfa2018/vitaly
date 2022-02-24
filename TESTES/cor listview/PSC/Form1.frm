VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Show Grid"
      Height          =   345
      Left            =   6180
      TabIndex        =   6
      Top             =   4620
      Width           =   2475
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008000&
      Caption         =   "three and more Color"
      Height          =   555
      Left            =   6090
      TabIndex        =   5
      Top             =   1470
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Font of ListView "
      Height          =   585
      Left            =   6120
      TabIndex        =   4
      Top             =   3960
      Width           =   2505
   End
   Begin VB.PictureBox picBg 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   510
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   " set equal space to diffrent color"
      Height          =   555
      Left            =   6060
      TabIndex        =   2
      Top             =   150
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Set checked to diffrent color"
      Height          =   525
      Left            =   6090
      TabIndex        =   1
      Top             =   780
      Width           =   2475
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4905
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   8652
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastCmd As Integer
'****************************************************************************
' *  Author: Deming Shang
' *  Email:  Ch21st@hotmail.com
' *  Desc.:  set any Listitem of a listview to any color
' *  Page:   http://www.msmvps.com/ch21st
' *
' *  Parameters      Type     Description
' *  ---------------------------------------------------------------------------
' *
' *  Return Value
' *  success return 0,else return error code
' *
' *
' *    Name          Date      Reason
' * -----------    ---------  -------------------------------------------------
' * Deming Shang   2003/11/24  New
' *
' ****************************************************************************'

Private Sub Check1_Click()
  lv.GridLines = Check1.Value

End Sub

Private Sub Command1_Click()
    Dim i As Integer
    LastCmd = 1
    picBg.Width = lv.Width
    picBg.Height = lv.ListItems(1).Height * (lv.ListItems.Count)
    picBg.ScaleHeight = lv.ListItems.Count
    picBg.ScaleWidth = 1
    picBg.DrawWidth = 1
    picBg.Cls
    For i = 1 To lv.ListItems.Count
    
       If lv.ListItems(i).Checked = True Then
         picBg.Line (0, i - 1)-(1, i), &HC0FFFF, BF
       Else
         picBg.Line (0, i - 1)-(1, i), &HFF8080, BF
       End If
    Next
    

    lv.Picture = picBg.Image

End Sub

Private Sub Command2_Click()
    Dim i As Integer
    LastCmd = 2
    picBg.Width = lv.Width
    picBg.Height = lv.ListItems(1).Height * (lv.ListItems.Count)
    picBg.ScaleHeight = lv.ListItems.Count
    picBg.ScaleWidth = 1
    picBg.DrawWidth = 1
    picBg.Cls

    For i = 1 To lv.ListItems.Count
       If i Mod 2 = 0 Then
         picBg.Line (0, i - 1)-(1, i), RGB(254, 209, 199), BF
       Else
         picBg.Line (0, i - 1)-(1, i), RGB(200, 125, 68), BF
       End If
    Next
    

    lv.Picture = picBg.Image

End Sub

Private Sub Command3_Click()
  lv.Font.Size = 15
  Select Case LastCmd
     Case 1
       Command1_Click
     Case 2
       Command2_Click
  End Select
     
End Sub

Private Sub Command4_Click()
    Dim i As Integer
    LastCmd = 1
    picBg.Width = lv.Width
    picBg.Height = lv.ListItems(1).Height * (lv.ListItems.Count)
    picBg.ScaleHeight = lv.ListItems.Count
    picBg.ScaleWidth = 1
    picBg.DrawWidth = 1
    picBg.Cls
    For i = 1 To lv.ListItems.Count
      Select Case i Mod 5
         Case 0
            picBg.Line (0, i - 1)-(1, i), &HC0E0FF, BF
         Case 1
            picBg.Line (0, i - 1)-(1, i), &HC0FFC0, BF
         Case 2
            picBg.Line (0, i - 1)-(1, i), &HFFC0FF, BF
         Case 3
            picBg.Line (0, i - 1)-(1, i), &HE0E0E0, BF
         Case 4
            picBg.Line (0, i - 1)-(1, i), &H8000&, BF
      End Select

    Next
    

    lv.Picture = picBg.Image

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mRow As ListItem
    
    Me.ScaleMode = vbTwips
  '---------------------------
  'initialize Listview Control
  '--------------------------
    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Checkboxes = True
    lv.ColumnHeaders.Add , , "ID"
    lv.ColumnHeaders.Add , , "Note"
    For i = 0 To 40
      Set mRow = lv.ListItems.Add(, , CStr(i))
      mRow.SubItems(1) = "This is Item " & i
    Next
  
   lv.ListItems(3).Checked = True
   lv.ListItems(5).Checked = True
   lv.ListItems(13).Checked = True
   lv.ListItems(23).Checked = True
   lv.ListItems(6).Checked = True
   lv.ListItems(9).Checked = True
   
  '-------------------------
  'initialize Picture style
  '------------------------
    picBg.BackColor = lv.BackColor
    picBg.ScaleMode = vbTwips
    picBg.BorderStyle = vbBSNone
    picBg.AutoRedraw = True
    picBg.Visible = False
   '---------------------------
   

End Sub
