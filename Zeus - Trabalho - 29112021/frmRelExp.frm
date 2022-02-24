VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmRelExp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de relatórios de expedição"
   ClientHeight    =   9090
   ClientLeft      =   420
   ClientTop       =   330
   ClientWidth     =   21450
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleMode       =   0  'User
   ScaleWidth      =   21450
   Begin VB.CommandButton cmdCadastro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   720
      Picture         =   "frmRelExp.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmRelExp.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "Salvar Relatório"
      ToolTipText     =   "Salvar Relatório"
      Top             =   8400
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Itens disponíveis para emissão do relatório"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   7200
      TabIndex        =   37
      Top             =   120
      Width           =   14175
      Begin VB.TextBox txtLvw 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8880
         TabIndex        =   43
         Top             =   7800
         Width           =   1000
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6975
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   12303
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483635
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
      Begin VB.TextBox txtCadastro 
         Height          =   330
         Index           =   17
         Left            =   1680
         TabIndex        =   22
         Top             =   7680
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRelExp.frx":265E
         TabIndex        =   46
         Top             =   7800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "frmRelExp.frx":26D6
         TabIndex        =   38
         Top             =   7320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmRelExp.frx":2736
         TabIndex        =   39
         Top             =   7320
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRelExp.frx":2796
         TabIndex        =   40
         Top             =   7320
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRelExp.frx":2804
         TabIndex        =   41
         Top             =   7320
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frmRelExp.frx":287E
         TabIndex        =   42
         Top             =   7320
         Width           =   2535
      End
      Begin IpToolTips.cIpToolTips cIpToolTips1 
         Left            =   11280
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         BackColor       =   0
      End
      Begin VB.Label Label9 
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   44
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Transporte "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   6975
      Begin ZEUS.chameleonButton chameleonButton2 
         Height          =   255
         Left            =   6480
         TabIndex        =   61
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   2
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRelExp.frx":2906
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   3360
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
      Begin VB.ComboBox cboCadastro 
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         ItemData        =   "frmRelExp.frx":2922
         Left            =   6120
         List            =   "frmRelExp.frx":2977
         TabIndex        =   16
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   11
         Left            =   5880
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   13
         Left            =   3120
         TabIndex        =   15
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox txtCadastro 
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "frmRelExp.frx":29E7
         TabIndex        =   55
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   6120
         OleObjectBlob   =   "frmRelExp.frx":2A47
         TabIndex        =   54
         Top             =   2040
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRelExp.frx":2AA5
         TabIndex        =   53
         Top             =   2040
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":2B0B
         TabIndex        =   52
         Top             =   2040
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":2B71
         TabIndex        =   51
         Top             =   1440
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "frmRelExp.frx":2BDB
         TabIndex        =   50
         Top             =   840
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":2C51
         TabIndex        =   49
         Top             =   840
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmRelExp.frx":2CB3
         TabIndex        =   48
         Top             =   240
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":2D15
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Veículo - Dados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   6735
         Begin VB.TextBox txtCadastro 
            Height          =   330
            Index           =   16
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   6495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRelExp.frx":2D7B
            TabIndex        =   60
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtCadastro 
            Height          =   330
            Index           =   15
            Left            =   2520
            TabIndex        =   19
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtCadastro 
            Height          =   330
            Index           =   14
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   2
            ItemData        =   "frmRelExp.frx":2DE7
            Left            =   3960
            List            =   "frmRelExp.frx":2E3C
            TabIndex        =   20
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   1
            ItemData        =   "frmRelExp.frx":2EAC
            Left            =   1560
            List            =   "frmRelExp.frx":2F01
            TabIndex        =   18
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   3960
            OleObjectBlob   =   "frmRelExp.frx":2F71
            TabIndex        =   59
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   2520
            OleObjectBlob   =   "frmRelExp.frx":2FD5
            TabIndex        =   58
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "frmRelExp.frx":304F
            TabIndex        =   57
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRelExp.frx":30B3
            TabIndex        =   56
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Relatório nº: "
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
      TabIndex        =   34
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Relatório "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   28
      Top             =   960
      Width           =   6975
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":312B
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":3199
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":31FF
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmRelExp.frx":326D
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmRelExp.frx":32D5
         TabIndex        =   33
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Data: "
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
      Left            =   2400
      TabIndex        =   27
      Top             =   120
      Width           =   1815
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   155516929
         CurrentDate     =   40449
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo do Movimento"
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
      Left            =   4320
      TabIndex        =   26
      Top             =   120
      Width           =   2775
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExp.frx":3341
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   495
      Left            =   7320
      OleObjectBlob   =   "frmRelExp.frx":33AB
      TabIndex        =   45
      Top             =   8400
      Width           =   14055
   End
End
Attribute VB_Name = "frmRelExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Abaixo - usado poder editar o listview --------------------
'straight from the standard API Viewver
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Declare Function GetScrollInfo Lib "user32" (ByVal HWnd As Long, ByVal fnBar As Long, lpScrollInfo As SCROLLINFO) As Long
 
'interestingly, API Viewer doesn't have these constants, translating from Windows.h is straight forward
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
  
'my declarations
Private Const c_EntryTxt = ""
Private m_ColIndex As Long 'listview col index
Private m_RowIndex As Long 'listview row index
'---------------------------------------------------

'Abaixo ajusta automaticamente a largura das colunas
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
'Acima ajusta automaticamente a largura das colunas
Private X As Integer, W As Integer
Private ContaLV As Integer, LinhaLV As Integer, ContaChecado As Integer, LimiTador As Integer
Private rsLocal As New ADODB.Recordset
Private vPonte1 As TextBox

Private Sub chameleonButton2_Click()
    ChamaGridTrans
    CarregaTipoTrans
    txtCadastro(14).SetFocus
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 1
    Case 3
    Case 4
        If ContaChecado > LimiTador Then
            mobjMsg.Abrir "Limite máximo de itens selecionados foi ultrapassado." & vbCrLf & "Limite Máximo: " & LimiTador, Ok, critico, "Atenção"
        Else
            mobjMsg.Abrir "Deseja gravar o relatório?", YesNo, pergunta, "Zeus"
            If Tp = 1 Then
                GravarDados
                Unload Me
            End If
        End If
    Case 6
        mobjMsg.Abrir "Deseja sair da tela de emissão de relatórios?", YesNo, pergunta, "Zeus"
        If Tp = 1 Then
            Unload Me
        End If
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Me.Top = 0
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    
    ContaChecado = 0
    LimiTador = 1000 '49
    Legenda = "Aguarde"
    SelecionaLinha
    CompoeControles
    listview_cabecalho 'Chama a Sub que monta o cabeçalho das colunas do Listview
    
    CompoeListview2 'Listview de Expedição
    
    txtLvw = ""
    'txtLvw.Visible = False
    txtLvw.Tag = False 'is ListView2 dirty, not used in this example
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delas e e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Posição", ListView1.Width / 13
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Desenho", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Rev.", ListView1.Width / 22
    ListView1.ColumnHeaders.Add , , "Q. Total", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Peso", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Q. Pendente", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Q. à lib.", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "codFase", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "UN", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "CodLM", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "CodSeq", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Peso Lib.", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Insp. Realizadas", ListView1.Width / 10000
'    ListView1.ColumnHeaders.Add , , "Possui Pintura?", ListView1.Width / 10000
    
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(13).Alignment = lvwColumnRight
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim rsRelInsp As New ADODB.Recordset
    Dim sqlRelInsp As String
    
    sqlRelInsp = "select a.codprojeto,a.fce,a.projeto,c.nome from tbProjetos as a inner join tbFO as b on a.fce = b.fce inner join tbclifor as c on b.codclifor = c.codclifor " & _
                 "where a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and a.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' Order by a.fce desc,a.descricao"
    rsRelInsp.Open sqlRelInsp, cnBanco, adOpenKeyset, adLockReadOnly
    If rsRelInsp.RecordCount = 0 Then Exit Sub
    
    txtCadastro(4) = Format(GeraCodigo, "000000000") & "" 'Identificador do relatório
    txtCadastro(0) = Val(Mid(varGlobal, 7, 4)) 'FCE nº
    txtCadastro(1) = rsRelInsp.Fields(0) 'ID Projeto
    txtCadastro(2) = rsRelInsp.Fields(2) 'Descrição do projeto
    txtCadastro(3) = rsRelInsp.Fields(3) 'Nome do cliente
    DTPicker1 = Date 'Data de emissão do relatório
    varGlobal2 = txtCadastro(4).Text
    SkinLabel7.Caption = vSituacao
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

Private Sub CompoeListview2()
On Error GoTo Err
    Dim rsLisview As New ADODB.Recordset
    Dim sqlLisview As String
    
    Dim rsLisview2 As New ADODB.Recordset
    Dim sqlLisview2 As String
    Dim W As Integer
    sqlLisview = "Select a.fce,a.codprojeto,a.codlm,a.codseq,max(a.desenho) as desenho,max(a.revisao) as revisao,max(a.descposicao) as descposicao,max(a.posicao) as posicao,sum(a.pesolib) as pesolib,max(a.status) as status,sum(a.qtdlib) as qtddisp  " & _
                 "from tbrelinspexpitens as a where a.status = 10 and a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and a.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' group by a.fce,a.codprojeto,a.codlm,a.codseq"
    rsLisview.Open sqlLisview, cnBanco, adOpenKeyset, adLockReadOnly
    If rsLisview.RecordCount <> 0 Then Principal.ProgressBar1.Max = rsLisview.RecordCount
    W = 0
    ListView1.ListItems.Clear
    Do While Not rsLisview.EOF
        Principal.ProgressBar1.Value = W
        sqlLisview2 = "Select a.fce,a.codprojeto,a.codlm,a.codseq,max(a.desenho) as desenho,max(a.revisao) as revisao,max(a.descposicao) as descposicao,max(a.posicao) as posicao,sum(a.pesolib) as pesolib,max(a.status) as status,sum(a.qtdlib) as qtddisp  " & _
                     "from tbrelinspexpitens as a where a.status = 11 and a.fce = '" & Val(Mid(varGlobal, 7, 4)) & "' and a.codprojeto = '" & Val(Mid(varGlobal, 1, 6)) & "' and a.codlm = '" & rsLisview.Fields(2) & "' and a.codseq ='" & rsLisview.Fields(3) & "' group by a.fce,a.codprojeto,a.codlm,a.codseq "
        rsLisview2.Open sqlLisview2, cnBanco, adOpenKeyset, adLockReadOnly
        
        If rsLisview2.RecordCount > 0 Then
            'ENTRA NESSA CONDIÇÃO SE O RESULTADO DA CONSULTA rsLisview2 NÃO FOR VAZIA
            'E SOMENTE ENTRA NA CONDIÇÃO ABAIXO SE A POSIÇÃO 10 DA PRIMEIRA CONSULTA - A POSIÇÃO 10 DA SEGUNDA CONSULTA FOR MAIOR QUE ZERO
            If rsLisview.Fields(10) - rsLisview2.Fields(10) > 0 Then
                Set ItemLst = ListView1.ListItems.Add(, , rsLisview.Fields(7)) 'Posição identificador
                ItemLst.SubItems(1) = rsLisview.Fields(6) 'Posição descrição
                ItemLst.SubItems(2) = rsLisview.Fields(4) 'Desenho
                ItemLst.SubItems(3) = rsLisview.Fields(5) 'Revisão
                ItemLst.SubItems(4) = rsLisview.Fields(10) 'Quantidade total
                ItemLst.SubItems(5) = Format(rsLisview.Fields(8), "#,##0.00;(#,##0.00)") 'Peso Posição
                
                ItemLst.SubItems(6) = rsLisview.Fields(10) - rsLisview2.Fields(10) 'Q. Pendente
                
                ItemLst.SubItems(7) = " " 'Quantidade à liberar
                ItemLst.SubItems(8) = "11" 'Código da Fase de (Expedição)
                ItemLst.SubItems(9) = "-" 'Unidade de medida
                ItemLst.SubItems(10) = rsLisview.Fields(2) 'Código da LM - Lista de Materiais
                ItemLst.SubItems(11) = rsLisview.Fields(3) 'Código da sequência da LM
                ItemLst.SubItems(12) = " " 'Peso Liberado da Posição
                ItemLst.SubItems(13) = "-" 'Inspeções realizadas
            End If
        Else
            'ENTRA NESSA CONDIÇÃO SE O RESULTADO DA CONSULTA rsLisview2 FOR VAZIA
            Set ItemLst = ListView1.ListItems.Add(, , rsLisview.Fields(7)) 'Posição identificador
            ItemLst.SubItems(1) = rsLisview.Fields(6) 'Posição descrição
            ItemLst.SubItems(2) = rsLisview.Fields(4) 'Desenho
            ItemLst.SubItems(3) = rsLisview.Fields(5) 'Revisão
            ItemLst.SubItems(4) = rsLisview.Fields(10) 'Quantidade total
            ItemLst.SubItems(5) = Format(rsLisview.Fields(8), "#,##0.00;(#,##0.00)") 'Peso Posição
            ItemLst.SubItems(6) = rsLisview.Fields(10) 'Q. Pendente
            ItemLst.SubItems(7) = " " 'Quantidade à liberar
            ItemLst.SubItems(8) = "11" 'Código da Fase de (Expedição)
            ItemLst.SubItems(9) = "-" 'Unidade de medida
            ItemLst.SubItems(10) = rsLisview.Fields(2) 'Código da LM - Lista de Materiais
            ItemLst.SubItems(11) = rsLisview.Fields(3) 'Código da sequência da LM
            ItemLst.SubItems(12) = " " 'Peso Liberado da Posição
            ItemLst.SubItems(13) = "-" 'Inspeções realizadas
        End If
        rsLisview2.Close
        W = W + 1
        If Not rsLisview.EOF Then rsLisview.MoveNext
    Loop
    Principal.ProgressBar1.Value = 0
    rsLisview.Close
    Set rsLisview = Nothing
    Set rsLisview2 = Nothing
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

Private Sub PosLinha()
    Dim ContaLV As Integer
    ContaLV = ListView1.ListItems.Count
    For LinhaLV = 1 To ContaLV
        If ListView1.ListItems.Item(LinhaLV).Selected = True Then
            Exit For
        End If
    Next
End Sub

Private Function SomaTotais()
On Error GoTo TrataErro
    SkinLabel12.Caption = ""
    SomaTotais = True
    Dim Y As Integer, SomaPeso As Double, SomaQtd As Double
    Y = ListView1.ListItems.Count
    SomaQtd = 0
    SomaPeso = 0
    For W = 1 To Y
        If ListView1.ListItems.Item(W).Checked = True Then
            ListView1.ListItems(W).Selected = True
            SomaQtd = SomaQtd + ListView1.SelectedItem.ListSubItems.Item(7)
            SomaPeso = SomaPeso + ((ListView1.SelectedItem.ListSubItems.Item(5) / ListView1.SelectedItem.ListSubItems.Item(4)) * ListView1.SelectedItem.ListSubItems.Item(7))
            ListView1.SelectedItem.ListSubItems.Item(12) = Format(((ListView1.SelectedItem.ListSubItems.Item(5) / ListView1.SelectedItem.ListSubItems.Item(4)) * ListView1.SelectedItem.ListSubItems.Item(7)), "#,##0.00;(#,##0.00)")
        End If
    Next
    SkinLabel29 = Format(SomaQtd, "#,##0.00;(#,##0.00)")
    SkinLabel30 = Format(SomaPeso, "#,##0.00;(#,##0.00)")
    txtCadastro(17) = Format(SomaPeso, "#,##0.00;(#,##0.00)")
    Exit Function
TrataErro:
    SomaTotais = False
    'mobjMsg.Abrir "Existem itens marcados no Relatorio que não possuem Qtd. liberada", Ok, informacao, "Atenção"
    SkinLabel12.Caption = "Existem itens marcados no Relatorio que não possuem Qtd. liberada"
End Function

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tbRelInspExp order by codrel Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraCodigo.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        If IniciaRelsEm > 0 Then
            GeraCodigo = IniciaRelsEm
        Else
            GeraCodigo = 1 'NovoCodigo
        End If
    End If
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Private Sub ChamaGridTrans()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD,a.INATIVO from " & vBancoTotvs & ".dbo.ttra as a order by a.nome"
    procnom = "nome"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Transportadoras"
    'Pesquisa = frmRelExp.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        Pesquisa = Mid$(Pesquisa, 7, 100)
        rsLocal.Find "nome=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadastro(6).Text = Format(rsLocal.Fields(0), "000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
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

Private Sub CarregaTipoTrans()
On Error GoTo Err
    Dim X As Integer
    Dim rsTipoTrans As New ADODB.Recordset
    SqlM = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD from " & vBancoTotvs & ".dbo.ttra as a order by a.CODETD"
    rsTipoTrans.Open SqlM, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsTipoTrans.EOF Then rsTipoTrans.MoveFirst
    rsTipoTrans.Find "CODTRA=" & "'" & Format(txtCadastro(6), "000") & "'"
    If rsTipoTrans.EOF Then
        txtCadastro(6).Text = Format(txtCadastro(6), "000") & ""
        If Val(Pesquisa) <> 0 Then
            Msgbox "Transportadora não cadastrada", vbInformation, "Zeus"
            txtCadastro(7) = ""
        End If
    Else
        txtCadastro(6).Text = Format(rsTipoTrans.Fields(0), "000") & "" 'codigo
        txtCadastro(7).Text = rsTipoTrans.Fields(1) 'nome
        If Not IsNull(rsTipoTrans.Fields(2)) Then txtCadastro(8).Text = rsTipoTrans.Fields(2) 'cnpj
        If Not IsNull(rsTipoTrans.Fields(3)) Then txtCadastro(9).Text = rsTipoTrans.Fields(3) 'ie
        If Not IsNull(rsTipoTrans.Fields(4)) Then txtCadastro(10).Text = rsTipoTrans.Fields(4) 'endereco (rua+numero)
        If Not IsNull(rsTipoTrans.Fields(5)) Then txtCadastro(11).Text = rsTipoTrans.Fields(5) 'cep
        If Not IsNull(rsTipoTrans.Fields(6)) Then txtCadastro(12).Text = rsTipoTrans.Fields(6) 'bairro
        If Not IsNull(rsTipoTrans.Fields(7)) Then txtCadastro(13).Text = rsTipoTrans.Fields(7) 'cidade
        If Not IsNull(rsTipoTrans.Fields(8)) Then cboCadastro(0).Text = rsTipoTrans.Fields(8) 'UF
        For X = 7 To 13
            txtCadastro(X).Enabled = False
        Next
        cboCadastro(0).Enabled = False
    End If
    rsTipoTrans.Close
    Set rsTipoTrans = Nothing
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

Private Sub GravarDados()
On Error GoTo Err

    Dim Y As Integer, X As Integer
    
    Dim rsRelatorio As New ADODB.Recordset
    Dim sqlRelatorio As String
    Dim rsItensRelatorio As New ADODB.Recordset
    Dim sqlItensRelatorio As String
    
    
    'If SomaTotais = False Then Exit Sub
    If Val(SkinLabel29) = 0 Then
        mobjMsg.Abrir "Nenhum Item Selecionado para compor o relatório", Ok, critico, "Atenção"
        Exit Sub
    End If
    
10  cnBanco.BeginTrans

    sqlRelatorio = "select * from tbRelInspExp"
    rsRelatorio.Open sqlRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    rsRelatorio.AddNew
    
    txtCadastro(4) = Format(GeraCodigo, "000000000") & "" 'Identificador do relatório
    
    rsRelatorio.Fields(0) = Val(txtCadastro(4)) 'Codigo do Relatorio
    rsRelatorio.Fields(1) = Val(txtCadastro(0)) 'FCE
    rsRelatorio.Fields(2) = Val(txtCadastro(1)) 'Codigo do projeto
    rsRelatorio.Fields(3) = Format(DTPicker1, "dd/mm/yyyy") 'Data do relatorio
    rsRelatorio.Fields(4) = txtCadastro(5) 'Observação
    rsRelatorio.Fields(5) = 0 'Status de impressão
    'rsRelatorio.Fields(6) = cboCadastro(4) 'Norma de Liberação
    rsRelatorio.Fields(7) = 11 'Tipo relatorio (11) Expedição
    rsRelatorio.Fields(8) = Format(txtCadastro(17), "#,##0.00;(#,##0.00)") 'Peso de balança
    rsRelatorio.Fields(9) = NomUsu
    rsRelatorio.Update
    rsRelatorio.Close
    Set rsRelatorio = Nothing

    'Gravar dados referente aos Itens do Relatório
    sqlItensRelatorio = "select * from tbRelInspExpitens"
    rsItensRelatorio.Open sqlItensRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If ListView1.ListItems.Item(X).Checked = True Then
            rsItensRelatorio.AddNew
            rsItensRelatorio.Fields(0) = Val(txtCadastro(4)) 'Codigo do relatorio
            rsItensRelatorio.Fields(1) = Val(txtCadastro(0).Text) 'Nº FCE
            rsItensRelatorio.Fields(2) = Val(txtCadastro(1)) 'Código do Projeto
            rsItensRelatorio.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2) 'Desenho
            rsItensRelatorio.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(3) 'Revisão do Desenho
            rsItensRelatorio.Fields(5) = ListView1.ListItems.Item(X) 'Posição
            rsItensRelatorio.Fields(6) = ListView1.SelectedItem.ListSubItems.Item(1) 'Descrição da posição
            rsItensRelatorio.Fields(7) = ListView1.SelectedItem.ListSubItems.Item(8) 'Status (Codfase)
            rsItensRelatorio.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(7) 'Quantidade liberada
            rsItensRelatorio.Fields(9) = ListView1.SelectedItem.ListSubItems.Item(12) 'Peso liberada
            rsItensRelatorio.Fields(10) = Val(ListView1.SelectedItem.ListSubItems.Item(10)) 'Código da LM - Lista de Material
            rsItensRelatorio.Fields(11) = Val(ListView1.SelectedItem.ListSubItems.Item(11)) 'Código da sequencia da LM
            rsItensRelatorio.Fields(13) = "-" 'Inspeções realizadas
            rsItensRelatorio.Update
        End If
    Next
    rsItensRelatorio.Close
    Set rsItemRelatorio = Nothing
    
    cnBanco.CommitTrans
    
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(20, 1) = txtCadastro(0).Text 'grava o numero da FCE
    vQualquerDado(1, 1) = txtCadastro(4).Text
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = txtCadastro(6).Text
    vQualquerDado(2, 2) = "I"
    vQualquerDado(3, 1) = txtCadastro(7).Text
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtCadastro(8).Text
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtCadastro(9).Text
    vQualquerDado(5, 2) = "S"
    
    vQualquerDado(6, 1) = txtCadastro(10).Text
    vQualquerDado(6, 2) = "S"
    vQualquerDado(7, 1) = txtCadastro(11).Text
    vQualquerDado(7, 2) = "S"
    vQualquerDado(8, 1) = txtCadastro(12).Text
    vQualquerDado(8, 2) = "S"
    vQualquerDado(9, 1) = txtCadastro(13).Text
    vQualquerDado(9, 2) = "S"
    vQualquerDado(10, 1) = cboCadastro(0).Text
    vQualquerDado(10, 2) = "S"
    
    vQualquerDado(11, 1) = txtCadastro(14).Text
    vQualquerDado(11, 2) = "S"
    vQualquerDado(12, 1) = cboCadastro(1).Text
    vQualquerDado(12, 2) = "S"
    vQualquerDado(13, 1) = txtCadastro(15).Text
    vQualquerDado(13, 2) = "S"
    vQualquerDado(14, 1) = cboCadastro(2).Text
    vQualquerDado(14, 2) = "S"
    vQualquerDado(15, 1) = txtCadastro(16).Text
    vQualquerDado(15, 2) = "S"
    GravaDados "tbRelExpTransp", "codrel", "I", txtCadastro(4), 15, "", "", txtCadastro(4)
   
    mobjMsg.Abrir "Dados gravados com sucesso", Ok, informacao, "Atenção"
    
    mobjMsg.Abrir "Dados gravados com sucesso.Deseja imprimir de relatório de Expedição?", YesNo, pergunta, "Zeus"
    If Tp = 1 Then
        vCodRel = Val(txtCadastro(4))
        FCRExpedicao.Show 1
        sqlRelatorio = "update tbRelInspExp set statusimp=1 where codrel ='" & vCodRel & "'"
        rsRelatorio.Open sqlRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    End If
    
    Unload Me
    
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub Form_Resize()
    DimensionaFormExp chamaForm
End Sub

Private Sub ListView1_Click()
    Dim M As Integer
    M = ListView1.ListItems.Count
    PosLinha
    ContaChecado = 0
    SkinLabel10 = "Registros selecionadas: "
    'compoeSiglas
    For K = 1 To M
        If ListView1.ListItems.Item(K).Selected = True Then
            If ListView1.ListItems.Item(K).Checked = False Then
                ListView1.ListItems.Item(K).Checked = True
                ContaChecado = ContaChecado + 1
                ListView1.SelectedItem.ListSubItems.Item(13) = "-"
                If ListView1.SelectedItem.ListSubItems.Item(7) = " " And txtLvw = "" Then
                    ListView1.SelectedItem.ListSubItems.Item(7) = ListView1.SelectedItem.ListSubItems.Item(6)
                End If
                If Val(ListView1.SelectedItem.ListSubItems.Item(7)) > Val(ListView1.SelectedItem.ListSubItems.Item(6)) Then ListView1.SelectedItem.ListSubItems.Item(7) = ListView1.SelectedItem.ListSubItems.Item(6)
            Else
                ListView1.ListItems.Item(K).Checked = False
                'ListView1.SelectedItem.ListSubItems.Item(7) = " "
                If txtLvw = "" Or ListView1.SelectedItem.ListSubItems.Item(7) = " " Then
                    ListView1.SelectedItem.ListSubItems.Item(12) = " "
                    ListView1.SelectedItem.ListSubItems.Item(13) = "-"
                    txtLvw = ""
                End If
            End If
        Else
            If ListView1.ListItems.Item(K).Checked = True Then
                ContaChecado = ContaChecado + 1
            End If
        End If
    Next
    SomaTotais
    SkinLabel10 = SkinLabel10 & ContaChecado
    If ListView1.ListItems.Count > 0 Then
        ListView1.ListItems(LinhaLV).Selected = True
    End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'fired when the listitem is already selected, for this reason can't used mousedown event
'so we know which row is clicked, for the column, we need to translate the x to listview coordinate
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 8 To 8
            PosLinha
            'If ListView1.ListItems.Item(LinhaLV).Checked = False Then Exit Sub
            
            leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView1.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView1.SelectedItem.Text
                    Else
                        .Text = ListView1.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
                Exit For
            End If
        Next i
    End If
End If
End Sub

Function GetLvwDeltaX() As Single
'returns deltaX, the scroll distance in pixels relative to ListView2.left, how much we scroll right
'si.npage propotional to both the width of the scroll box and ListView2.width
'si.npos is the scrolling position, which is propotional to deltaX

    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView1.ColumnHeaders(ListView1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    'PrintLvwColInfo
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView1.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    '58 is some constant to get things just right
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
'called from ListView2 mouseup and subclass scroll events
'constants used are determined by trial & error, these are mainly the various widths and heights
'of edges in the classical windows. these constants may not be correct for other windows styles.
Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
If m_ColIndex Then
    If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
    Set lvwCol = ListView1.ColumnHeaders(m_ColIndex)
        
    txtLeft = ListView1.Left + lvwCol.Left + 48 - dx
    If txtLeft < ListView1.Left Then txtLeft = ListView1.Left + 48
    
    txtRightMax = ListView1.Left + ListView1.Width - 48
    If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
    If m_ColIndex = ListView1.ColumnHeaders.Count Then
        txtRight = txtRightMax
    Else
        txtRight = ListView1.Left + ListView1.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
        If txtRight > txtRightMax Then txtRight = txtRightMax
    End If
    
    txtWidth = txtRight - txtLeft
    If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
    txtTopMin = ListView1.Top
    If Not ListView1.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
    txtTopMax = ListView1.Top + ListView1.Height
    If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
    txtTop = ListView1.Top + ListView1.SelectedItem.Top + 54
    If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
    With txtLvw '.move produces runtime error with -ve values
        .Left = 10960
        .Top = txtTop
        .Width = txtWidth
        .Height = ListView1.SelectedItem.Height - 8
    End With
End If
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    If Index = 6 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaTipoTrans
        End If
    End If
Error:
    Exit Sub
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView2 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
            If ListView1.SelectedItem.ListSubItems.Item(7) > ListView1.SelectedItem.ListSubItems.Item(6) Then ListView1.SelectedItem.ListSubItems.Item(7) = ListView1.SelectedItem.ListSubItems.Item(6)
            If ListView1.ListItems.Item(LinhaLV).Checked = False Then ListView1.SelectedItem.ListSubItems.Item(7) = ""
        'other keys can be used for navigation
    End Select
End Sub

Private Sub txtLvw_LostFocus()
    If m_ColIndex = 1 Then
        ListView1.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
    ElseIf m_ColIndex Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView2's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

'FUNCAO PARA MUDAR TOOLTIPS
Private Sub MudaTool()
    On Error Resume Next
    Dim Ctl As Control
    Dim i As Integer
    With Me.cIpToolTips1
        .Create
        .Title = "Atenção:" 'Titulo do tooltip
        .MyIcon = itInfoIcon 'Icone do tooltip
        .BackColor = &H80000018  'Cor de fundo
        .ForeColor = &H800000    'Cor da letra e bordas
        For Each Ctl In Me.Controls
            If Ctl.Tag <> "" Then
                .AddTool Ctl, tfAbsolute, Replace(Ctl.Tag, "|", vbCrLf)
            End If
        Next
    End With
End Sub

Private Sub CompoeTabTemp()
On Error GoTo Err
    Dim rsCodRel As New ADODB.Recordset
    Dim rsTbTemp As New ADODB.Recordset
    Dim sqlCodRel As String, sqlTbTemp As String
    Dim Y As Integer, X As Integer
    
    sqlTbTemp = "Delete from tbtemp"
    rsTbTemp.Open sqlTbTemp, cnBanco
    
    sqlTbTemp = "Select * from tbtemp"
    rsTbTemp.Open sqlTbTemp, cnBanco, adOpenKeyset, adLockOptimistic
    
    Y = frmRelInsp.ListView1.ListItems.Count
    For X = 1 To Y
        frmRelInsp.ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If frmRelInsp.ListView1.ListItems.Item(X).Checked = True Then
            sqlCodRel = "select tbrelatorios.codrel,tbitemrelatorio.idld from tbitemrelatorio inner join tbRelatorios on tbrelatorios.codrel = tbitemrelatorio.codrel where tbitemrelatorio.idld = '" & Val(frmRelInsp.ListView1.ListItems.Item(X)) & "'" & " order by tbitemrelatorio.codprocesso,tbitemrelatorio.codfase desc"
            rsCodRel.Open sqlCodRel, cnBanco, adOpenKeyset, adLockReadOnly
            rsCodRel.MoveNext
            If rsCodRel.RecordCount > 0 Then
                rsTbTemp.AddNew
                rsTbTemp.Fields(0) = rsCodRel.Fields(0)
                rsTbTemp.Fields(1) = rsCodRel.Fields(1)
            End If
            rsCodRel.Close
        End If
    Next
    Set rsCodRel = Nothing
    rsTbTemp.Update
    Set rsTbTemp = Nothing
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

Private Sub SelecionaLinha()
    Dim Y As Integer
    Y = vListViewPrincipal.ListItems.Count
    For W = 1 To Y
        If vListViewPrincipal.ListItems.Item(W).Selected = True Then
            Exit For
        End If
    Next
    vListViewPrincipal.ListItems(W).Selected = True
End Sub

