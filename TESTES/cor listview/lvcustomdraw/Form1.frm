VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "VB ListView Custom Draw demo"
   ClientHeight    =   4200
   ClientLeft      =   2235
   ClientTop       =   1770
   ClientWidth     =   6360
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6360
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
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
'
' Brad Martinez http://www.mvps.org
'
' Demonstrates how to do Custom Draw with the VB ListView.
'
' ========================================================
' This project uses subclassing, and utilizes the services of the "Debug
' Object for AddressOf Subclassing" ActiveX server, Dbgwproc.dll, which
' allows unencumbered code execution when stepping through code in
' the VB IDE. This server is freely distributable and can be obtained from
' Microsoft at http://msdn.microsoft.com/vbasic/downloads/controls.asp.

' Set the conditional compilation argument:   DEBUGWINDOWPROC = 1
' in the project properties dialog/Make tab to enable the server's services.
' ========================================================

Private Sub Form_Load()
  Dim i As Long
  Dim Item As ListItem
  
  Randomize
  
  ' Set the Form's ScaleMode to pixels for column resizing in Form_Resize
  ScaleMode = vbPixels
  
  ' Load up the global color array with the 16 basic colors
  'For i = 0 To 15: g_crl16(i) = QBColor(i): Next
    
  ' Initialize the ListView
  With ListView1
    .View = lvwReport
    .Font.Size = 10
    .MultiSelect = True
    
    For i = 1 To 4
      .ColumnHeaders.Add , , "column" & i
    Next
    
    For i = 0 To &H3F
      Set Item = .ListItems.Add(, "L" & i, "item" & i)
      Item.SubItems(1) = i * 10
      Item.SubItems(2) = i * 100
      Item.SubItems(3) = i * 1000
      Item.Tag = QBColor(CInt(Rnd * 15)) & " " & QBColor(CInt(Rnd * 15))
    Next
  End With

  ' Subclass the Form so we can process the NM_CUSTOMDRAW
  ' notification message sent from the ListView.
  SubClassLV hWnd, AddressOf WndProcLV, ListView1
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call UnSubClassLV(hWnd)
  'Erase g_crl16
End Sub

Private Sub mnuCDNewDraw_Click()
  g_fNewDraw = Not g_fNewDraw
  ListView1.Refresh
End Sub

Private Sub Timer1_Timer()
Dim Item As ListItem
    For Each Item In ListView1.ListItems
        Item.SubItems(1) = CStr(CInt(Item.SubItems(1)) + 1)
    Next
End Sub
