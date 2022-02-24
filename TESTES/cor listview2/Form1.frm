VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7646
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



Private Sub Command1_Click()
MudaCorList ListView1, 1
End Sub

Private Sub Form_Load()
  Dim Lis As ListItem
  Dim X As Double
 
  Picture1.BackColor = ListView1.BackColor
  Picture1.ScaleMode = vbTwips
  Picture1.BorderStyle = vbBSNone
  Picture1.AutoRedraw = True
  Picture1.Visible = False
 
  ListView1.ColumnHeaders.Add , , "COLUNA1"
  ListView1.ColumnHeaders.Add , , "COLUNA2"
  ListView1.ColumnHeaders.Add , , "COLUNA3"
  ListView1.ColumnHeaders.Add , , "COLUNA4"
  ListView1.View = lvwReport
 
  For X = 1 To 50
    Set Lis = ListView1.ListItems.Add
    With Lis
      .Text = ""
      .SubItems(1) = 12615935
      If X > 5 Then .SubItems(1) = &H8000000F
  
      .SubItems(2) = "LINHA" & X
      .SubItems(3) = "LINHA" & X
    End With
  Next
End Sub
Public Function MudaCorList(Listview As Listview, Coluna As Integer)
    Dim i As Integer
    'LastCmd = 1
    Picture1.Width = Listview.Width
    Picture1.Height = Listview.ListItems(1).Height * (Listview.ListItems.Count)
    Picture1.ScaleHeight = Listview.ListItems.Count
    Picture1.ScaleWidth = 1
    Picture1.DrawWidth = 1
    Picture1.Cls
    For i = 1 To Listview.ListItems.Count
   
       If Trim(Listview.ListItems.Item(i).SubItems(Coluna)) <> "" Then
         Picture1.Line (0, i - 1)-(1, i), Trim(Listview.ListItems.Item(i).SubItems(Coluna)), BF
       Else
         Picture1.Line (0, i - 1)-(1, i), &HFEFAE0, BF
       End If
    Next
   
    Listview.Picture = Picture1.Image
    ListView1.Refresh
End Function
