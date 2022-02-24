VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2400
      List            =   "Form1.frx":0019
      TabIndex        =   1
      Text            =   "Código"
      Top             =   360
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0062
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":01BC
            Key             =   "cima"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0316
            Key             =   "baixo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0470
            Key             =   "raio"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05CA
            Key             =   "bandeira"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0724
            Key             =   "email_novo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":087E
            Key             =   "email_lido"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09D8
            Key             =   "clip"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B32
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DA1
            Key             =   "nao"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Treinamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Origem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Introdutório"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Obrigatório"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ativo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    listview_cabecalho
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 16, , "raio"
    ListView1.ColumnHeaders.Add , , "Treinamento", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Origem", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Introdutório", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Obrigatório", ListView1.Width / 15
    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 11
    ListView1.ColumnHeaders.Add , , "Ativo", ListView1.Width / 15
    Compoe_Listview

'    ListView1.Sorted = True
'    ListView1.SortOrder = lvwAscending
'    ListView1.SortKey = 5

End Sub

Private Sub Compoe_Listview()
    
    'Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    'ListView1.ListItems.Clear 'Limpa o listview
    'Set ItemLst = ListView1.ListItems.Add(, , Format(rsListview(0), "000000"))
    'ItemLst.SubItems(1) = "" & rsListview.Fields(1)
    'ItemLst.SubItems(2) = "" & rsListview.Fields(3)
    
    Dim lstEntry As ListItem
    Set lstEntry = ListView1.ListItems.Add(, , "000001", "email_novo")
    lstEntry.SubItems(1) = "Treinamento-01"
    lstEntry.SubItems(2) = "Interno"
    lstEntry.SubItems(3) = ""
    lstEntry.SubItems(4) = ""
    lstEntry.SubItems(5) = "Funcional"
    lstEntry.SubItems(6) = ""
    lstEntry.ListSubItems.Item(3).ReportIcon = "ok"
    lstEntry.ListSubItems.Item(4).ReportIcon = "ok"
    lstEntry.ListSubItems.Item(6).ReportIcon = "ok"
    ListView1.Refresh
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Combo1.Text = ColumnHeader.Text
End Sub
