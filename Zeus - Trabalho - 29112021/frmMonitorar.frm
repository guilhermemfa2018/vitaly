VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmMonitorar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoramento da Produção"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19515
   Icon            =   "frmMonitorar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   19515
   Begin VB.CommandButton Command1 
      Caption         =   "Encerrar apropriação"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   30
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":0CCA
            Key             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":435C
            Key             =   "FC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":79EE
            Key             =   "P"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":B080
            Key             =   "F"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":E712
            Key             =   "I"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitorar.frx":11DA4
            Key             =   "O"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame12 
      Caption         =   "Setor Selecionado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4695
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmMonitorar.frx":15436
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações da Apropriação "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   4815
      Begin VB.Frame Frame9 
         Caption         =   "Dados da Tarefa "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Width           =   4575
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":1548E
            TabIndex        =   26
            Top             =   1680
            Width           =   3975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmMonitorar.frx":154E8
            TabIndex        =   25
            Top             =   1080
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   1920
            OleObjectBlob   =   "frmMonitorar.frx":15542
            TabIndex        =   24
            Top             =   1080
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmMonitorar.frx":1559C
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":155F6
            TabIndex        =   20
            Top             =   1080
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":15650
            TabIndex        =   19
            Top             =   1440
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmMonitorar.frx":156C4
            TabIndex        =   18
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   1920
            OleObjectBlob   =   "frmMonitorar.frx":1572E
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmMonitorar.frx":1578E
            TabIndex        =   16
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":157F6
            TabIndex        =   15
            Top             =   840
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":15854
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":158B4
            TabIndex        =   34
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Apropriado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   10
         Top             =   1800
         Width           =   2415
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":1590E
            TabIndex        =   28
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Orçado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   2415
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMonitorar.frx":15968
            TabIndex        =   27
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Sub-centro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   2415
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":159C2
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Satatus "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   4575
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   495
            Left            =   240
            OleObjectBlob   =   "frmMonitorar.frx":15A1C
            TabIndex        =   21
            Top             =   840
            Width           =   4215
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            ForeColor       =   &H80000008&
            Height          =   530
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   4305
            TabIndex        =   14
            Top             =   240
            Width           =   4335
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage7 
               Height          =   480
               Left            =   2400
               ToolTipText     =   "Situação definida pelo DP"
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":15A76
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage6 
               Height          =   480
               Left            =   1920
               ToolTipText     =   "Apropriando sem registrar ponto"
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":1910C
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage5 
               Height          =   480
               Left            =   1440
               ToolTipText     =   "Não está na empresa"
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":1C7A2
               Enabled         =   0   'False
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage4 
               Height          =   480
               Left            =   0
               ToolTipText     =   "Apropriando"
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":1FE38
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage3 
               Height          =   480
               Left            =   480
               ToolTipText     =   "Apropriando em parada"
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":234CE
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
               Height          =   480
               Left            =   960
               ToolTipText     =   "Registrou ponto. Porém, não está apropriando"
               Top             =   0
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   847
               Image           =   "frmMonitorar.frx":26B64
               Props           =   5
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Colaborador "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         Begin VB.PictureBox Picture1 
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   1755
            TabIndex        =   12
            Top             =   240
            Width           =   1815
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   2055
               Left            =   -120
               Top             =   0
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   3625
               Image           =   "frmMonitorar.frx":2A1FA
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colaboradores (Status)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   14295
      Begin MSComctlLib.ListView ListView3 
         Height          =   8175
         Left            =   9720
         TabIndex        =   29
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   14420
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   8415
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   14843
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   873
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fábrica (Setores)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4695
      Begin MSComctlLib.ListView ListView2 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Label Label53 
      Height          =   255
      Left            =   5160
      TabIndex        =   31
      Top             =   9360
      Width           =   4695
   End
End
Attribute VB_Name = "frmMonitorar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const LVM_FIRST = &H1000
    Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
    Private Const LVSCW_AUTOSIZE = -1
    Private Const LVSCW_AUTOSIZE_USEHEADER = -2


Private vStatus As String
Private vSubCentro As String
Private vChapaEncerra As String
Private vPosition As Integer

Private Sub Command1_Click()
    EncerraAprop
End Sub

Private Sub Form_Load()
    listview_cabecalho
    CompoeLV
    HabBotao 1
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_Resize()
    DimensionaForm
'    Dim XSize As Integer
'    Dim YSize As Integer
    
'    On Error Resume Next
'    If Form.WindowState <> 0 Then Exit Sub
    
'    Me.Top = 0
'    Me.Left = 0
'    Me.Height = Me.Height * YSize
'    Me.Width = Me.Width * XSize
    
'    For i = 0 To Me.Controls.Count - 1
'        Me.Controls(i).Left = Me.Controls(i).Left * XSize
'        Me.Controls(i).Top = Me.Controls(i).Top * YSize
'        Me.Controls(i).Height = Me.Controls(i).Height * YSize
'        Me.Controls(i).Width = Me.Controls(i).Width * XSize
'    Next i
End Sub

Private Sub ListView1_Click()
    CompoeControles
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    CompoeControles
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    CompoeControles
End Sub

Private Sub ListView2_Click()
    If ListView2.ListItems.Item(1).Selected = True Then
        SkinLabel20 = "PREPARAÇÃO"
        vSubCentro = "'3000.3101.SC-01','3000.3101.SC-02','3000.3101.SC-03','3000.3101.SC-04','3000.3101.SC-05','3000.3101.SC-06','3000.3101.SC-07','3000.3101.SC-08','3000.3101.SC-09','3000.3101.SC-10','3000.3101.SC-12','3000.3102.SC-01','3000.3102.SC-02','3000.3106.SC-01'"
    End If
    If ListView2.ListItems.Item(2).Selected = True Then
        SkinLabel20 = "MONTAGEM"
        vSubCentro = "'3000.3103.SC-01','3000.3103.SC-02'"
    End If
    If ListView2.ListItems.Item(3).Selected = True Then
        SkinLabel20 = "SOLDA"
        vSubCentro = "'3000.3104.SC-01','3000.3104.SC-02'"
    End If
    If ListView2.ListItems.Item(4).Selected = True Then
        SkinLabel20 = "ACABAMENTO"
        vSubCentro = "'3000.3105.SC-01','3000.3105.SC-02','3000.3105.SC-05','3000.3105.SC-06'"
    End If
    If ListView2.ListItems.Item(5).Selected = True Then
        SkinLabel20 = "MANUTENÇÃO"
        vSubCentro = "'4000.4101.SC-01','4000.4101.SC-02','4001.4101.AJ-01'"
    End If
    
    If ListView2.ListItems.Item(6).Selected = True Then
        SkinLabel20 = "USINAGEM"
        vSubCentro = "'4001.4101.SC-01','4001.4101.SC-02','4001.4101.SC-03','4001.4101.SC-04','4001.4101.SC-05','4001.4101.SC-06','4001.4101.SC-07','4001.4101.SC-08','4001.4101.SC-09','4001.4001.SC-11'"
    End If
    CompoeLV1
    RedimensionaColuna
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Colaborador", ListView1.Width / 1.3
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Centro Custo", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "CC", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "OS", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "OS Rev.", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Operação", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Grupo", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Orçamento", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "C.Barra", ListView1.Width / 10000
    
    ListView1.ColumnHeaders.Add , , "P. Entrada", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "P. Saida", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "  ", ListView1.Width / 14
    
    ListView1.ColumnHeaders.Add , , "R. Entrada", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "R. Saida", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Situação", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "FCE", ListView1.Width / 10000
'    ListView1.View = lvwList 'Modo de Exibição do seu Listview
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview

    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "C.Barra", ListView3.Width / 3
    ListView3.ColumnHeaders.Add , , "Entrada", ListView3.Width / 4
    ListView3.ColumnHeaders.Add , , "Saida", ListView3.Width / 4
    ListView3.ColumnHeaders.Add , , "Parada", ListView3.Width / 4
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview

    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , " ", ListView4.Width / 1.291
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "", ListView4.Width / 10000

    ListView4.ColumnHeaders.Add , , "Relógio de Ponto", ListView4.Width / 4
    ListView4.ColumnHeaders.Add , , "  ", ListView4.Width / 14
    
    ListView4.ColumnHeaders.Add , , "Parametrizado S. Pessoal", ListView4.Width / 4
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    
    Me.ListView4.ColumnHeaders(12).Alignment = lvwColumnCenter
    Me.ListView4.ColumnHeaders(13).Alignment = lvwColumnCenter


End Sub

Private Sub CompoeLV1()
On Error GoTo Err
    Dim rsStatus As New ADODB.Recordset
    Dim sqlStatus As String
'    sqlStatus = "select b.chapa,b.NOME,a.codigobarra,CONVERT (VARCHAR, a.dataent,103) as dataent,CONVERT (VARCHAR, a.horaent, 108) as horaent,CONVERT (VARCHAR, a.datasai,103) as datasai,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada,f.NOME,f.CODREDUZIDO,case when a.codigobarra in('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') then 'FC' when a.codigobarra Is null or a.codigobarra in('9019','9020') then 'P'  Else 'A' end Status," & _
'                "c.idos,c.revisaoos,c.idoperacao,c.grupo,dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido " & _
'                "from CORPORERM.dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA and a.dataent = CONVERT (date, GETUTCDATE()) and a.horasai is null left join CORPORERM.dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join CORPORERM.dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra " & _
'                "where b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO <> 'D' or b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO = 'D' AND GETDATE ( )<b.DTDESLIGAMENTO+1 Order by f.CODREDUZIDO,b.NOME"
    
    
'    sqlStatus = "Select b.chapa,b.NOME,a.codigobarra,CONVERT (VARCHAR, a.dataent,103) as dataent,CONVERT (VARCHAR, a.horaent, 108) as horaent,CONVERT (VARCHAR, a.datasai,103) as datasai,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada,f.NOME,f.CODREDUZIDO,case when a.codigobarra in('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') then 'FC' when a.codigobarra Is null or a.codigobarra in('9019','9020') then 'P'  Else 'A' end Status,c.idos,c.revisaoos,c.idoperacao,c.grupo,dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido " & _
'                "from " & vBancoTotvs & ".dbo.PFUNC as b left join tbOsMov as a on a.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = b.CHAPA and a.dataent = CONVERT (date, GETUTCDATE()) and a.horasai is null left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra " & _
'                "where b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO <> 'D' or b.CODSITUACAO in('A','D','F') and b.CHAPA > 5 and f.CODREDUZIDO in (" & vSubCentro & ") and b.CODSITUACAO = 'D' AND GETDATE ( )<b.DTDESLIGAMENTO+1 " & _
'                "union " & _
'                "select b.chapa COLLATE SQL_Latin1_General_CP1_CI_AI,b.nome COLLATE SQL_Latin1_General_CP1_CI_AI,a.codigobarra COLLATE SQL_Latin1_General_CP1_CI_AI,CONVERT (VARCHAR, a.dataent ,103) as dataent,CONVERT (VARCHAR, a.horaent, 108) as horaent,CONVERT (VARCHAR, a.datasai,103) as datasai,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada COLLATE SQL_Latin1_General_CP1_CI_AI,b.nmcc COLLATE SQL_Latin1_General_CP1_CI_AI,b.idcc COLLATE SQL_Latin1_General_CP1_CI_AI,case when a.codigobarra in('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') then 'FC' " & _
'                "when a.codigobarra Is null or a.codigobarra in('9019','9020') then 'P'  Else 'A' end Status,c.idos,c.revisaoos,c.idoperacao,c.grupo,dbo.FN_CONVMIN(cast(replace(replace(c.tempocalc,'.',''),',','.') as money)) as Tempo_Convertido " & _
'                "from tbTerceirizados as b left join tbOsMov as a on a.chapa = b.CHAPA and a.dataent = CONVERT (date, GETUTCDATE()) and a.horasai is null left join " & vBancoTotvs & ".dbo.PFRATEIOFIXO as e on b.CHAPA COLLATE SQL_Latin1_General_CP1_CI_AS = e.CHAPA left join " & vBancoTotvs & ".dbo.GCCUSTO as f on e.CODCCUSTO = f.CODCCUSTO left join tbMPItens as c on a.codigobarra = c.codigobarra where b.ativo = 'S' and b.idcc in (" & vSubCentro & ") Order by f.CODREDUZIDO,b.NOME"
    

'    sqlStatus = sqlStatus & "SELECT " & vbCrLf
'    sqlStatus = sqlStatus & " B.CHAPA, " & vbCrLf
'    sqlStatus = sqlStatus & " B.NOME, " & vbCrLf
'    sqlStatus = sqlStatus & " A.CODIGOBARRA, " & vbCrLf
'    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.DATAENT,103) AS DATAENT, " & vbCrLf
'    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.HORAENT, 108) AS HORAENT, " & vbCrLf
'    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.DATASAI,103) AS DATASAI, " & vbCrLf
'    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.HORASAI, 108) AS HORASAI, " & vbCrLf
'    sqlStatus = sqlStatus & " A.IDPARADA, " & vbCrLf
'    sqlStatus = sqlStatus & " F.NOME, " & vbCrLf
'    sqlStatus = sqlStatus & " F.CODREDUZIDO, " & vbCrLf
'    sqlStatus = sqlStatus & " CASE " & vbCrLf
'    sqlStatus = sqlStatus & "     WHEN A.CODIGOBARRA IN('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') THEN 'FC' " & vbCrLf
'    sqlStatus = sqlStatus & "     WHEN A.CODIGOBARRA IS NULL OR A.CODIGOBARRA IN('9019','9020') THEN 'P'  ELSE 'A' " & vbCrLf
'    sqlStatus = sqlStatus & " END STATUS, " & vbCrLf
'    sqlStatus = sqlStatus & " C.IDOS, " & vbCrLf
'    sqlStatus = sqlStatus & " C.REVISAOOS, " & vbCrLf
'    sqlStatus = sqlStatus & " C.IDOPERACAO, " & vbCrLf
'    sqlStatus = sqlStatus & " C.GRUPO, " & vbCrLf
'    sqlStatus = sqlStatus & " DBO.FN_CONVMIN(CAST(REPLACE(REPLACE(C.TEMPOCALC,'.',''),',','.') AS MONEY)) AS " & vbCrLf
'    sqlStatus = sqlStatus & " TEMPO_CONVERTIDO " & vbCrLf
'    sqlStatus = sqlStatus & "FROM CORPORERM.DBO.PFUNC AS B " & vbCrLf
'    sqlStatus = sqlStatus & "LEFT JOIN TBOSMOV AS A ON " & vbCrLf
'    sqlStatus = sqlStatus & " A.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS = B.CHAPA AND " & vbCrLf
'    sqlStatus = sqlStatus & " A.DATAENT = CONVERT (DATE, GETUTCDATE()) AND " & vbCrLf
'    sqlStatus = sqlStatus & " A.HORASAI IS NULL " & vbCrLf
'    sqlStatus = sqlStatus & "LEFT JOIN CORPORERM.DBO.PFRATEIOFIXO AS E ON " & vbCrLf
'    sqlStatus = sqlStatus & " B.CHAPA = E.CHAPA AND " & vbCrLf
'    sqlStatus = sqlStatus & " B.CODCOLIGADA = E.CODCOLIGADA " & vbCrLf
'    sqlStatus = sqlStatus & "LEFT JOIN CORPORERM.DBO.GCCUSTO AS F ON " & vbCrLf
'    sqlStatus = sqlStatus & " E.CODCCUSTO = F.CODCCUSTO AND " & vbCrLf
'    sqlStatus = sqlStatus & " B.CODCOLIGADA = F.CODCOLIGADA " & vbCrLf
'    sqlStatus = sqlStatus & "LEFT JOIN TBMPITENS AS C ON " & vbCrLf
'    sqlStatus = sqlStatus & " A.CODIGOBARRA = C.CODIGOBARRA " & vbCrLf
'    sqlStatus = sqlStatus & "WHERE " & vbCrLf
'    sqlStatus = sqlStatus & " F.CODREDUZIDO IN (" & vSubCentro & ") AND " & vbCrLf
'    sqlStatus = sqlStatus & " B.CODCOLIGADA = 6 AND " & vbCrLf
'    sqlStatus = sqlStatus & " B.CHAPA > 5 AND " & vbCrLf
'    sqlStatus = sqlStatus & " ( " & vbCrLf
'    sqlStatus = sqlStatus & "     B.CODSITUACAO <> 'D' OR " & vbCrLf
'    sqlStatus = sqlStatus & "     B.CODSITUACAO IN('A','F') OR " & vbCrLf
'    sqlStatus = sqlStatus & "     B.CODSITUACAO = 'D' AND GETDATE ( )<B.DTDESLIGAMENTO+1 " & vbCrLf
'    sqlStatus = sqlStatus & " ) " & vbCrLf
'    sqlStatus = sqlStatus & " " & vbCrLf
'    sqlStatus = sqlStatus & " " & vbCrLf
    
    sqlStatus = sqlStatus & "SELECT " & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA, " & vbCrLf
    sqlStatus = sqlStatus & " B.NOME, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(A.CODIGOBARRA,'-') AS CODIGOBARRA, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(CONVERT(VARCHAR,A.DATAENT,103),'-') AS DATAENT, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(CONVERT(VARCHAR,A.HORAENT,108),'-') AS HORAENT, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(CONVERT(VARCHAR,A.DATASAI,103),'-') AS DATASAI, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(CONVERT(VARCHAR,A.HORASAI,108),'-') AS HORASAI, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(A.IDPARADA,'-') AS IDPARADA, " & vbCrLf
    sqlStatus = sqlStatus & " F.NOME, " & vbCrLf
    sqlStatus = sqlStatus & " F.CODREDUZIDO, " & vbCrLf
    sqlStatus = sqlStatus & " CASE " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN B.CODSITUACAO <> 'A' THEN " & vbCrLf
    sqlStatus = sqlStatus & "         'O' " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN G.BATIDA2 IS NOT NULL AND G.BATDA3 IS NULL THEN" & vbCrLf
    sqlStatus = sqlStatus & "         'F' " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN G.BATIDA1 IS NULL AND A.HORAENT IS NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "         'F' /*NÃO HÁ REGISTRO DE PONTO E NÃO HÁ REGISTRO DE APROPRIACAO*/ " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN G.BATIDA1 IS NULL AND A.HORAENT IS NOT NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "         'I' /*APROPRIANDO SEM REGISTRAR PONTO*/ " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN G.BATIDA1 IS NOT NULL AND A.HORAENT IS NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "         'P' /*REGISTROU PONTO, MAS NÃO ESTÁ APROPRIANDO*/ " & vbCrLf
    sqlStatus = sqlStatus & "     ELSE " & vbCrLf
    sqlStatus = sqlStatus & "         CASE " & vbCrLf
    sqlStatus = sqlStatus & "             WHEN A.CODIGOBARRA IN('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') THEN 'FC' /*OCIOSO - APROPRIANDO EM PARADA*/ " & vbCrLf
    sqlStatus = sqlStatus & "             WHEN A.CODIGOBARRA IS NULL OR A.CODIGOBARRA IN('9019','9020') THEN 'P' " & vbCrLf
    sqlStatus = sqlStatus & "             ELSE 'A' " & vbCrLf
    sqlStatus = sqlStatus & "         END " & vbCrLf
    sqlStatus = sqlStatus & " END STATUS, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(CONVERT(VARCHAR,C.IDOS),'-') AS IDOS, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(C.REVISAOOS,'-') AS REVISAOOS, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(CONVERT(VARCHAR,C.IDOPERACAO),'-') AS IDOPERACAO, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(C.GRUPO,'-') AS GRUPO, " & vbCrLf
    sqlStatus = sqlStatus & " ISNULL(DBO.FN_CONVMIN(CAST(REPLACE(REPLACE(C.TEMPOCALC,'.',''),',','.') AS  MONEY)),'-') AS TEMPO_CONVERTIDO, " & vbCrLf
    
    sqlStatus = sqlStatus & " ENTRADA = " & vbCrLf
    sqlStatus = sqlStatus & "     CASE " & vbCrLf
    sqlStatus = sqlStatus & "         WHEN G.BATIDA1 IS NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "             '-' " & vbCrLf
    sqlStatus = sqlStatus & "         Else " & vbCrLf
    sqlStatus = sqlStatus & "             CASE " & vbCrLf
    sqlStatus = sqlStatus & "                WHEN G.BATDA3 IS NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "                    CONVERT(VARCHAR(5),G.BATIDA1) " & vbCrLf
    sqlStatus = sqlStatus & "                Else " & vbCrLf
    sqlStatus = sqlStatus & "                    CONVERT(VARCHAR(5),G.BATDA3) " & vbCrLf
    sqlStatus = sqlStatus & "            End " & vbCrLf
    sqlStatus = sqlStatus & "     END, " & vbCrLf
    sqlStatus = sqlStatus & " SAIDA = " & vbCrLf
    sqlStatus = sqlStatus & "     CASE " & vbCrLf
    sqlStatus = sqlStatus & "         WHEN G.BATIDA2 IS NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "             '-' " & vbCrLf
    sqlStatus = sqlStatus & "         Else " & vbCrLf
    sqlStatus = sqlStatus & "             CASE " & vbCrLf
    sqlStatus = sqlStatus & "                WHEN G.BATDA3 IS NULL THEN " & vbCrLf
    sqlStatus = sqlStatus & "                    CONVERT(VARCHAR(5),G.BATIDA2) " & vbCrLf
    sqlStatus = sqlStatus & "                Else " & vbCrLf
    sqlStatus = sqlStatus & "                    ISNULL(CONVERT(VARCHAR(5),G.BATIDA4),'-') " & vbCrLf
    sqlStatus = sqlStatus & "            End " & vbCrLf
    sqlStatus = sqlStatus & "     End, ENTRADA_PARAM = ISNULL(CONVERT(VARCHAR(5),H.HORARIO_ENTRADA,108),'-'), SAIDA_PARAM = ISNULL(CONVERT(VARCHAR(5),H.HORARIO_SAIDA,108),'-'), B.CODSITUACAO, I.DESCRICAO, J.FCE " & vbCrLf
    
    
    
    'sqlStatus = sqlStatus & " ENTRADA = " & vbCrLf
    'sqlStatus = sqlStatus & "     CASE " & vbCrLf
    'sqlStatus = sqlStatus & "         WHEN G.BATIDA1 IS NULL THEN " & vbCrLf
    'sqlStatus = sqlStatus & "             '-' " & vbCrLf
    'sqlStatus = sqlStatus & "         ELSE " & vbCrLf
    'sqlStatus = sqlStatus & "             CONVERT(VARCHAR(5),G.BATIDA1) " & vbCrLf
    'sqlStatus = sqlStatus & "     END, " & vbCrLf
    'sqlStatus = sqlStatus & " SAIDA = " & vbCrLf
    'sqlStatus = sqlStatus & "     CASE " & vbCrLf
    'sqlStatus = sqlStatus & "         WHEN G.BATIDA2 IS NULL THEN " & vbCrLf
    'sqlStatus = sqlStatus & "             '-' " & vbCrLf
    'sqlStatus = sqlStatus & "         ELSE " & vbCrLf
    'sqlStatus = sqlStatus & "             CONVERT(VARCHAR(5),G.BATIDA2) " & vbCrLf
    'sqlStatus = sqlStatus & "     END, ENTRADA_PARAM = ISNULL(CONVERT(VARCHAR(5),H.HORARIO_ENTRADA,108),'-'), SAIDA_PARAM = ISNULL(CONVERT(VARCHAR(5),H.HORARIO_SAIDA,108),'-'), B.CODSITUACAO, I.DESCRICAO, J.FCE " & vbCrLf
    
    sqlStatus = sqlStatus & "FROM CORPORERM.DBO.PFUNC AS  B " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBOSMOV AS A ON " & vbCrLf
    sqlStatus = sqlStatus & " A.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS = B.CHAPA  AND " & vbCrLf
    sqlStatus = sqlStatus & " A.DATAENT = CONVERT (DATE,  GETUTCDATE()) AND " & vbCrLf
    sqlStatus = sqlStatus & " A.HORASAI IS NULL " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN CORPORERM.DBO.PFRATEIOFIXO AS  E ON " & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA = E.CHAPA   AND " & vbCrLf
    sqlStatus = sqlStatus & " B.CODCOLIGADA = E.CODCOLIGADA " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN CORPORERM.DBO.GCCUSTO AS F ON " & vbCrLf
    sqlStatus = sqlStatus & " E.CODCCUSTO = F.CODCCUSTO AND " & vbCrLf
    sqlStatus = sqlStatus & " B.CODCOLIGADA = F.CODCOLIGADA " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBMPITENS AS C ON " & vbCrLf
    sqlStatus = sqlStatus & " A.CODIGOBARRA = C.CODIGOBARRA " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBPONTO AS G ON " & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA = G.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS AND " & vbCrLf
    sqlStatus = sqlStatus & " G.DATABATIDA = CONVERT(DATE, GETUTCDATE()) " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBHORARIOS AS H ON" & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA = H.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS  " & vbCrLf
    sqlStatus = sqlStatus & "INNER JOIN CORPORERM.DBO.PCODSITUACAO AS I ON B.CODSITUACAO = I.CODINTERNO " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBDESENHOSOS AS J ON C.IDOS = J.IDOS " & vbCrLf
    sqlStatus = sqlStatus & "WHERE " & vbCrLf
    sqlStatus = sqlStatus & " F.CODREDUZIDO IN (" & vSubCentro & ") AND " & vbCrLf
    sqlStatus = sqlStatus & " B.CODCOLIGADA = 6 AND " & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA > 5 AND " & vbCrLf
    sqlStatus = sqlStatus & " ( " & vbCrLf
    sqlStatus = sqlStatus & "     B.CODSITUACAO <> 'D' OR " & vbCrLf
    sqlStatus = sqlStatus & "     B.CODSITUACAO IN('A','F','E','L','T','X','Y') OR " & vbCrLf
    sqlStatus = sqlStatus & "     B.CODSITUACAO = 'D' AND GETDATE ( )<B.DTDESLIGAMENTO+1 " & vbCrLf
    sqlStatus = sqlStatus & " )"
    sqlStatus = sqlStatus & " UNION " & vbCrLf
    sqlStatus = sqlStatus & " " & vbCrLf
    sqlStatus = sqlStatus & "SELECT " & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    sqlStatus = sqlStatus & " B.NOME COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    sqlStatus = sqlStatus & " A.CODIGOBARRA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.DATAENT ,103) AS DATAENT, " & vbCrLf
    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.HORAENT, 108) AS HORAENT, " & vbCrLf
    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.DATASAI,103) AS DATASAI, " & vbCrLf
    sqlStatus = sqlStatus & " CONVERT (VARCHAR, A.HORASAI, 108) AS HORASAI, " & vbCrLf
    sqlStatus = sqlStatus & " A.IDPARADA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    sqlStatus = sqlStatus & " B.NMCC COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    sqlStatus = sqlStatus & " B.IDCC COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    sqlStatus = sqlStatus & " CASE " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN A.CODIGOBARRA IN('9003','9004','9005','9006','9007','9008','9009','9010','9011','9012','9013','9014','9015','9016','9017') THEN 'FC' " & vbCrLf
    sqlStatus = sqlStatus & "     WHEN A.CODIGOBARRA IS NULL OR A.CODIGOBARRA IN('9019','9020') THEN 'P' " & vbCrLf
    sqlStatus = sqlStatus & "     ELSE 'A' " & vbCrLf
    sqlStatus = sqlStatus & " END STATUS, " & vbCrLf
    sqlStatus = sqlStatus & " C.IDOS, " & vbCrLf
    sqlStatus = sqlStatus & " C.REVISAOOS, " & vbCrLf
    sqlStatus = sqlStatus & " C.IDOPERACAO, " & vbCrLf
    sqlStatus = sqlStatus & " C.GRUPO, " & vbCrLf
    sqlStatus = sqlStatus & " DBO.FN_CONVMIN(CAST(REPLACE(REPLACE(C.TEMPOCALC,'.',''),',','.') AS MONEY)) AS TEMPO_CONVERTIDO, ENTRADA = '-', SAIDA = '-', ENTRADA_PARAM = '-', SAIDA_PARAM = '-', CODSITUACAO = 'A', DESCRICAO = 'Ativo', J.FCE " & vbCrLf
    sqlStatus = sqlStatus & "FROM TBTERCEIRIZADOS AS B " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBOSMOV AS A ON " & vbCrLf
    sqlStatus = sqlStatus & " A.CHAPA = B.CHAPA AND " & vbCrLf
    sqlStatus = sqlStatus & " A.DATAENT = CONVERT (DATE, GETUTCDATE()) AND " & vbCrLf
    sqlStatus = sqlStatus & " A.HORASAI IS NULL " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN CORPORERM.DBO.PFRATEIOFIXO AS E ON " & vbCrLf
    sqlStatus = sqlStatus & " B.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AS = E.CHAPA " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN CORPORERM.DBO.GCCUSTO AS F ON " & vbCrLf
    sqlStatus = sqlStatus & " E.CODCCUSTO = F.CODCCUSTO AND " & vbCrLf
    sqlStatus = sqlStatus & " E.CODCOLIGADA = F.CODCOLIGADA " & vbCrLf
    sqlStatus = sqlStatus & "LEFT JOIN TBMPITENS AS C ON " & vbCrLf
    sqlStatus = sqlStatus & " A.CODIGOBARRA = C.CODIGOBARRA LEFT JOIN TBDESENHOSOS AS J ON C.IDOS = J.IDOS " & vbCrLf
    sqlStatus = sqlStatus & "WHERE " & vbCrLf
    sqlStatus = sqlStatus & " B.ATIVO = 'S' AND " & vbCrLf
    sqlStatus = sqlStatus & " B.IDCC IN (" & vSubCentro & ") AND " & vbCrLf
    sqlStatus = sqlStatus & " (E.CODCOLIGADA = 6 OR E.CODCOLIGADA IS NULL) " & vbCrLf
    sqlStatus = sqlStatus & "ORDER BY F.CODREDUZIDO,B.NOME"
    rsStatus.Open sqlStatus, cnBanco, adOpenKeyset, adLockReadOnly
    Dim ItemLst As ListItem
    Dim x As Integer
    x = 0
    ListView1.ListItems.Clear
    While Not rsStatus.EOF
        If rsStatus.Fields(10) = "A" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 1)
        ElseIf rsStatus.Fields(10) = "FC" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 2)
        ElseIf rsStatus.Fields(10) = "P" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 3)
        ElseIf rsStatus.Fields(10) = "F" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 4)
        ElseIf rsStatus.Fields(10) = "I" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 5)
        ElseIf rsStatus.Fields(10) = "O" Then
            Set ItemLst = ListView1.ListItems.Add(, , rsStatus.Fields(0) & " - " & rsStatus.Fields(1) & " (" & Mid$(rsStatus.Fields(8), 19, 30) & ")", 1, 6)
        End If
        ItemLst.SubItems(1) = "" & rsStatus.Fields(0)
        ItemLst.SubItems(2) = "" & Mid$(rsStatus.Fields(8), 19, 30)
        ItemLst.SubItems(3) = "" & rsStatus.Fields(10)
        ItemLst.SubItems(4) = "" & Mid$(rsStatus.Fields(8), 1, 15)
        ItemLst.SubItems(5) = "" & rsStatus.Fields(11)
        ItemLst.SubItems(6) = "" & rsStatus.Fields(12)
        ItemLst.SubItems(7) = "" & rsStatus.Fields(13)
        ItemLst.SubItems(8) = "" & rsStatus.Fields(14)
        ItemLst.SubItems(9) = "" & rsStatus.Fields(15)
        ItemLst.SubItems(10) = "" & rsStatus.Fields(2)
        
        ItemLst.SubItems(11) = "" & rsStatus.Fields(16)
        ItemLst.SubItems(12) = "" & rsStatus.Fields(17)
        ItemLst.SubItems(14) = "" & rsStatus.Fields(18)
        ItemLst.SubItems(15) = "" & rsStatus.Fields(19)
        ItemLst.SubItems(16) = "" & rsStatus.Fields(21)
        ItemLst.SubItems(17) = "" & rsStatus.Fields(22)
        
        rsStatus.MoveNext
        x = x + 1
    Wend
    rsStatus.Close
    Set rsStatus = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
End Sub

Private Sub CompoeLV()
    Dim ItemLst As ListItem
    ListView2.ColumnHeaders.Add , , "", ListView2.Width / 1.1
    Set ItemLst = ListView2.ListItems.Add(, , "Preparação")
    Set ItemLst = ListView2.ListItems.Add(, , "Montagem")
    Set ItemLst = ListView2.ListItems.Add(, , "Solda")
    Set ItemLst = ListView2.ListItems.Add(, , "Acabamento")
    Set ItemLst = ListView2.ListItems.Add(, , "Manutenção")
    Set ItemLst = ListView2.ListItems.Add(, , "Usinagem")
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim mStream As ADODB.Stream
    
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Dim rsCompoe As New ADODB.Recordset
    Dim sqlCompoe As String
    
    Dim Y As Integer, x As Integer
    Y = ListView1.ListItems.Count
    For x = 1 To Y
        If ListView1.ListItems.Item(x).Selected = True Then
            vPosition = x
            Exit For
        End If
    Next
    If ListView1.SelectedItem.ListSubItems.Item(3) = "A" Then  'Verde
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiNoGrayScale
        aicAlphaImage5.Opacity = 50
        aicAlphaImage5.grayScale = aiCCIR709
        aicAlphaImage6.grayScale = aiCCIR709
        aicAlphaImage7.grayScale = aiCCIR709
        SkinLabel4.Caption = "APROPRIANDO"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "FC" Then 'Laranja
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiNoGrayScale
        aicAlphaImage4.grayScale = aiCCIR709
        aicAlphaImage5.Opacity = 50
        aicAlphaImage5.grayScale = aiCCIR709
        aicAlphaImage6.grayScale = aiCCIR709
        aicAlphaImage7.grayScale = aiCCIR709
        SkinLabel4.Caption = "OCIOSO"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "P" Then 'Vermelho
        aicAlphaImage2.grayScale = aiNoGrayScale
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiCCIR709
        aicAlphaImage5.Opacity = 50
        aicAlphaImage5.grayScale = aiCCIR709
        aicAlphaImage6.grayScale = aiCCIR709
        aicAlphaImage7.grayScale = aiCCIR709
        SkinLabel4.Caption = "PARADO"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "F" Then 'Preto
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiCCIR709
        aicAlphaImage5.Opacity = 100
        aicAlphaImage5.grayScale = aiNoGrayScale
        aicAlphaImage6.grayScale = aiCCIR709
        aicAlphaImage7.grayScale = aiCCIR709
        SkinLabel4.Caption = "NÃO ESTÁ NA EMPRESA"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "I" Then 'Amarelo
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiCCIR709
        aicAlphaImage5.Opacity = 50
        aicAlphaImage5.grayScale = aiCCIR709
        aicAlphaImage6.grayScale = aiNoGrayScale
        aicAlphaImage7.grayScale = aiCCIR709
        SkinLabel4.Caption = "APROPRIANDO S/ PONTO"
    ElseIf ListView1.SelectedItem.ListSubItems.Item(3) = "O" Then 'Azul
        aicAlphaImage2.grayScale = aiCCIR709
        aicAlphaImage3.grayScale = aiCCIR709
        aicAlphaImage4.grayScale = aiCCIR709
        aicAlphaImage5.Opacity = 50
        aicAlphaImage5.grayScale = aiCCIR709
        aicAlphaImage6.grayScale = aiCCIR709
        aicAlphaImage7.grayScale = aiNoGrayScale
        SkinLabel4.Caption = UCase(ListView1.SelectedItem.ListSubItems.Item(16))
    End If
    SkinLabel5.Caption = ListView1.SelectedItem.ListSubItems.Item(4)
    SkinLabel8.Caption = ListView1.SelectedItem.ListSubItems.Item(5)
    SkinLabel7.Caption = ListView1.SelectedItem.ListSubItems.Item(17)
    SkinLabel9.Caption = ListView1.SelectedItem.ListSubItems.Item(6)
    SkinLabel10.Caption = ListView1.SelectedItem.ListSubItems.Item(7)
    SkinLabel11.Caption = ListView1.SelectedItem.ListSubItems.Item(8)
    SkinLabel12.Caption = ListView1.SelectedItem.ListSubItems.Item(9)
    SkinLabel15.Caption = ListView1.SelectedItem.ListSubItems.Item(10)
    If ListView1.SelectedItem.ListSubItems.Item(3) <> "A" Then
        SkinLabel5.Caption = "-"
        SkinLabel8.Caption = "-"
        SkinLabel9.Caption = "-"
        SkinLabel10.Caption = "-"
        SkinLabel11.Caption = "-"
        SkinLabel12.Caption = "-"
        SkinLabel13.Caption = "-"
        SkinLabel7.Caption = "-"
        SkinLabel15.Caption = ListView1.SelectedItem.ListSubItems.Item(10)
        If SkinLabel15.Caption <> "" Then
            Dim rsAchaParada As New ADODB.Recordset
            Dim sqlAchaParada As String
            sqlAchaParada = "select a.nmparada from tbParadas as a where a.codigo = '" & SkinLabel15.Caption & "'"
            rsAchaParada.Open sqlAchaParada, cnBanco, adOpenKeyset, adLockReadOnly
            If rsAchaParada.RecordCount > 0 Then
                SkinLabel11 = rsAchaParada.Fields(0)
            End If
            rsAchaParada.Close
            Set rsAchaParada = Nothing
        End If
    End If
    
    'HABILITA BOTÃO PARA FINALIZAR APROPRIAÇÃO
    HabBotao x
    
    'PEGA IMAGEM GRAVADO NO BANCO SQL E EXIBE EM UM COMPONENTE DE IMAGEM
    If Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 5) <> "CONTR" Then
        sqlCompoe = "select c.IDIMAGEM,a.chapa,a.nome,b.IMAGEM from " & vBancoTotvs & ".dbo.PFUNC as a left join " & vBancoTotvs & ".dbo.PPESSOA as c on a.CODPESSOA = c.CODIGO left join " & vBancoTotvs & ".dbo.GIMAGEM as b on c.IDIMAGEM = b.ID " & _
                    "where a.CHAPA = '" & ListView1.SelectedItem.ListSubItems.Item(1) & "'  order by a.nome"
    Else
        sqlCompoe = "select a.foto,a.chapa,a.nome,a.foto from tbTerceirizados as a where a.CHAPA = '" & ListView1.SelectedItem.ListSubItems.Item(1) & "' order by a.nome"
    End If
    rsCompoe.Open sqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 5) <> "CONTR" Then
        Set mStream = New ADODB.Stream
        mStream.Type = adTypeBinary
        mStream.Open
        mStream.Write rsCompoe.Fields(3).Value
        mStream.SaveToFile App.Path & "\Temp.jpg", adSaveCreateOverWrite
        aicAlphaImage1.ClearImage
        aicAlphaImage1.LoadImage_FromFile (App.Path & "\temp.jpg")
        Kill App.Path & "\Temp.jpg"
    Else
        Label53 = rsCompoe.Fields(3) 'Local onde esta armazenado a foto do coloborador
        aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    End If
    
    rsCompoe.Close
    Set rsCompoe = Nothing
    calculaTempoApropriado
    If Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 5) <> "CONTR" Then
        compoeAprop Mid$(ListView1.ListItems.Item(x), 1, 5)
    Else
        compoeAprop Mid$(ListView1.ListItems.Item(x), 1, 11)
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub HabBotao(vPosicao As Integer)
On Error GoTo Err
    Dim rsFimAprop As New ADODB.Recordset
    Dim sqlFimAprop As String
    sqlFimAprop = "Select a.multiplic from tbusuarios as a where a.nome= '" & NomUsu & "'"
    rsFimAprop.Open sqlFimAprop, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFimAprop.Fields(0) = "S" Then
        Command1.Visible = True
        If aicAlphaImage4.grayScale = aiNoGrayScale Or aicAlphaImage3.grayScale = aiNoGrayScale Then
            Command1.Enabled = True
        Else
            Command1.Enabled = False
        End If
    End If
    vChapaEncerra = Mid$(ListView1.ListItems.Item(vPosicao), 1, 5)
    
    rsFimAprop.Close
    Set rsFimAprop = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Sub

Private Sub compoeAprop(vChapa As String)
On Error GoTo Err
    Dim rsAprop As New ADODB.Recordset
    Dim sqlAprop As String
    Dim ItemLst As ListItem
    
    sqlAprop = "select a.codigobarra,CONVERT (VARCHAR, a.horaent, 108) as entrada,CONVERT (VARCHAR, a.horasai, 108) as horasai,a.idparada from tbOsMov as a where a.dataent = CONVERT (date, GETUTCDATE()) and a.chapa = '" & vChapa & "' order by a.chapa,a.horaent"
    rsAprop.Open sqlAprop, cnBanco, adOpenKeyset, adLockReadOnly
    ListView3.ListItems.Clear
    While Not rsAprop.EOF
        Set ItemLst = ListView3.ListItems.Add(, , rsAprop.Fields(0))
        ItemLst.SubItems(1) = "" & rsAprop.Fields(1)
        ItemLst.SubItems(2) = "" & rsAprop.Fields(2)
        If rsAprop.Fields(2) <> "" And rsAprop.Fields(3) = "" Then
            ItemLst.SubItems(3) = "Baixa indevida"
            ItemLst.ListSubItems(3).ForeColor = &HC0&
        Else
            ItemLst.SubItems(3) = "" & rsAprop.Fields(3)
        End If
        rsAprop.MoveNext
    Wend
    rsAprop.Close
    Set rsAprop = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub calculaTempoApropriado()
On Error GoTo Err
    Dim rsHAprop As New ADODB.Recordset
    Dim sqlHAprop As String
    Dim vHorasApropriadas As String
    
    If ListView1.SelectedItem.ListSubItems.Item(10) <> "" Then
        sqlHAprop = "select CONVERT (VARCHAR, a.horasai-a.horaent, 108) as horaent from tbOsMov  as a where a.codigobarra = '" & ListView1.SelectedItem.ListSubItems.Item(10) & "'"
        rsHAprop.Open sqlHAprop, cnBanco, adOpenKeyset, adLockReadOnly
        vHorasApropriadas = "00:00"
        Do While Not rsHAprop.EOF
            If Not IsNull(rsHAprop.Fields(0)) Then somaTempoPPSAtraso rsHAprop.Fields(0), vHorasApropriadas
            rsHAprop.MoveNext
        Loop
        rsHAprop.Close
        Set rsHAprop = Nothing
    End If
    If SkinLabel12 = "-" Or SkinLabel12 = "" Then
        SkinLabel13 = "-"
    Else
        SkinLabel13 = vHorasApropriadas
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub EncerraAprop()
On Error GoTo Err
    Dim rsFimAprop As New ADODB.Recordset
    Dim sqlFimAprop As String
    Dim vID As String
    sqlFimAprop = "select id from tbOsMov where chapa = '" & vChapaEncerra & "' and datasai is null"
    rsFimAprop.Open sqlFimAprop, cnBanco, adOpenKeyset, adLockReadOnly
    If rsFimAprop.RecordCount > 0 Then
        vID = rsFimAprop.Fields(0)
        rsFimAprop.Close
        Set rsFimAprop = Nothing
        If Time > "17:00:00" Then ' Maior que 17:00 horas
        'If Time < "17:00:00" Then ' Menor que 17:00 horas
            sqlFimAprop = "update tbOsMov set horasai = '17:00:00', datasai = '" & Format(Date, "YYYY-MM-DD") & "', idparada = '9018' where id = '" & vID & "'"
            rsFimAprop.Open sqlFimAprop, cnBanco
            CompoeLV1
            compoeAprop vChapaEncerra
            ListView1.ListItems.Item(vPosition).Selected = True
            CompoeControles
        ElseIf Time > "11:00:00" And Time < "12:00:00" Then
            sqlFimAprop = "update tbOsMov set horasai = '11:00:00', datasai = '" & Format(Date, "YYYY-MM-DD") & "', idparada = '9018' where id = '" & vID & "'"
            rsFimAprop.Open sqlFimAprop, cnBanco
            CompoeLV1
            compoeAprop vChapaEncerra
            ListView1.ListItems.Item(vPosition).Selected = True
            CompoeControles
        Else
            mobjMsg.Abrir "Apropriação fora do período de encerramento", , critico
        End If
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Function somaTempoPPSAtraso(vTempo, vOndeAcumula As String)
    Dim seg As Long, min As Long, hora As Long
    Dim tempo As Long
    Dim matriz2

    matriz2 = Split(vTempo, ":")
    tempo = tempo + (CLng(matriz2(0)) * 3600)
    tempo = tempo + (CLng(matriz2(1)) * 60)
    
    If vOndeAcumula <> "" Then
        matriz2 = Split(vOndeAcumula, ":")
        tempo = tempo + (CLng(matriz2(0)) * 3600)
        tempo = tempo + (CLng(matriz2(1)) * 60)
    End If
    
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    
    vOndeAcumula = Format(hora, "0000") & ":" & Format(min, "00")
    somaTempoPPSAtraso = vOndeAcumula
End Function

Private Sub RedimensionaColuna()
    Dim Column As Long
    Dim Counter As Long
    Counter = 0
    
    'SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, 0, LVSCW_AUTOSIZE_USEHEADER
    ListView1.ColumnHeaders.Item(1).Width = ListView1.Width * 40 / 100
    SendMessage ListView1.HWnd, LVM_SETCOLUMNWIDTH, 11, LVSCW_AUTOSIZE_USEHEADER
    SendMessage ListView1.HWnd, LVM_SETCOLUMNWIDTH, 12, LVSCW_AUTOSIZE_USEHEADER
    SendMessage ListView1.HWnd, LVM_SETCOLUMNWIDTH, 14, LVSCW_AUTOSIZE_USEHEADER
    SendMessage ListView1.HWnd, LVM_SETCOLUMNWIDTH, 15, LVSCW_AUTOSIZE_USEHEADER
    
    ListView4.ColumnHeaders.Item(1).Width = ListView1.ColumnHeaders.Item(1).Width + 50
    ListView4.ColumnHeaders.Item(12).Width = ListView1.ColumnHeaders.Item(12).Width + ListView1.ColumnHeaders.Item(13).Width
    ListView4.ColumnHeaders.Item(14).Width = ListView1.ColumnHeaders.Item(15).Width + ListView1.ColumnHeaders.Item(16).Width
    
    'For Column = Counter To ListView1.ColumnHeaders.Count - 2
        'SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, Column, LVSCW_AUTOSIZE_USEHEADER
    'Next
End Sub
