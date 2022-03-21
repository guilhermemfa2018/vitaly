VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmImpostosServicos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impostos e Serviços"
   ClientHeight    =   10665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10560
   Icon            =   "frmImpostosServicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   3600
      Top             =   10080
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
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
      Left            =   9360
      TabIndex        =   26
      Top             =   9960
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   17
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdImpostoServico 
      Height          =   615
      Index           =   12
      Left            =   720
      Picture         =   "frmImpostosServicos.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Sair"
      Top             =   9960
      Width           =   615
   End
   Begin VB.CommandButton cmdImpostoServico 
      Height          =   615
      Index           =   11
      Left            =   120
      Picture         =   "frmImpostosServicos.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Salvar"
      Top             =   9960
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critérios "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9855
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   10335
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9600
         Top             =   6600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmImpostosServicos.frx":265E
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmImpostosServicos.frx":3070
               Key             =   "EXC"
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3975
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7011
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Fórmulas"
         TabPicture(0)   =   "frmImpostosServicos.frx":3A82
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Legenda"
         TabPicture(1)   =   "frmImpostosServicos.frx":3A9E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame7"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3495
            Left            =   -74880
            TabIndex        =   31
            Top             =   360
            Width           =   9855
            Begin VB.Frame Frame4 
               BackColor       =   &H00B7B7B7&
               Caption         =   "Variáveis "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3015
               Left            =   120
               TabIndex        =   36
               Top             =   360
               Visible         =   0   'False
               Width           =   9615
               Begin VB.TextBox Text6 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   2655
                  Left            =   120
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   37
                  Text            =   "frmImpostosServicos.frx":3ABA
                  ToolTipText     =   "Legenda referente aos objetos que podem ser utilizados na fórmula"
                  Top             =   240
                  Width           =   9375
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00B7B7B7&
               Caption         =   "Matrizes "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3015
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Visible         =   0   'False
               Width           =   9615
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   2655
                  Left            =   120
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   35
                  Text            =   "frmImpostosServicos.frx":465C
                  ToolTipText     =   "Legenda referente aos objetos que podem ser utilizados na fórmula"
                  Top             =   240
                  Width           =   9375
               End
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Matrizes"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1680
               TabIndex        =   33
               Top             =   120
               Width           =   1695
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Variáveis"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00B7B7B7&
            Caption         =   "Campos com fórmulas "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   9855
            Begin VB.OptionButton Option2 
               Caption         =   "Matrizes"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1680
               TabIndex        =   29
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Variáveis"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.TextBox txtImpostoServico 
               Height          =   345
               Index           =   5
               Left            =   4320
               TabIndex        =   21
               Tag             =   "Composição da Fórmula do IMPOSTO ou SERVIÇO"
               Top             =   2160
               Width           =   5415
            End
            Begin VB.TextBox txtImpostoServico 
               Height          =   345
               Index           =   3
               Left            =   4320
               TabIndex        =   20
               Tag             =   "Composição da Fórmula do IMPOSTO ou SERVIÇO"
               Top             =   1320
               Width           =   5415
            End
            Begin VB.TextBox txtImpostoServico 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   4
               Left            =   4320
               TabIndex        =   19
               Tag             =   "Percentual a ser Aplicado sobre o valor do orçamento"
               Top             =   480
               Width           =   5415
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "frmImpostosServicos.frx":4A63
               TabIndex        =   22
               Top             =   240
               Width           =   3015
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "frmImpostosServicos.frx":4AD3
               TabIndex        =   23
               Top             =   1080
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "frmImpostosServicos.frx":4B4F
               TabIndex        =   24
               Top             =   1920
               Width           =   3975
            End
            Begin MSComctlLib.ListView lstListView 
               Height          =   2655
               Left            =   120
               TabIndex        =   25
               Top             =   720
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   4683
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   8388608
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
            Begin MSComctlLib.ListView lstListView2 
               Height          =   2655
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   4683
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
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
         End
      End
      Begin VB.CommandButton cmdImpostoServico 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   1920
         Picture         =   "frmImpostosServicos.frx":4BD9
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Excluir"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdImpostoServico 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   1320
         Picture         =   "frmImpostosServicos.frx":58A3
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "Editar"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdImpostoServico 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   720
         Picture         =   "frmImpostosServicos.frx":656D
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Novo"
         Top             =   6480
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8160
         TabIndex        =   15
         ToolTipText     =   "Seleciono o Tipo: IMPOSTO ou SERVIÇO"
         Top             =   240
         Width           =   2055
         Begin VB.ComboBox cboImpostoServico 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmImpostosServicos.frx":7237
            Left            =   120
            List            =   "frmImpostosServicos.frx":7241
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "Selecione o Tipo: IMPOSTO ou SERVIÇO"
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdImpostoServico 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   120
         Picture         =   "frmImpostosServicos.frx":7257
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Inserir"
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox txtImpostoServico 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Tag             =   "Descrição do Imposto ou Serviço"
         Top             =   1200
         Width           =   10095
      End
      Begin VB.TextBox txtImpostoServico 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Tag             =   "Nome do imposto ou Serviço"
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtImpostoServico 
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
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Identificador do Imposto ou Serviço"
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmImpostosServicos.frx":7F21
         TabIndex        =   12
         Top             =   960
         Width           =   3735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmImpostosServicos.frx":7FC3
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmImpostosServicos.frx":8025
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         Tag             =   "Lista de IMPOSTOS e SERVIÇOS cadastrados"
         Top             =   7200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   8388608
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
   End
End
Attribute VB_Name = "frmImpostosServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IMPOSTOS E SERVIÇOS
Private vPonte1 As TextBox, vPonte2 As TextBox
Private rsCriterio As New ADODB.Recordset
Private SqlCriterio As String
Private vOndeEstaOTab As Integer

Private Sub cmdImpostoServico_Click(Index As Integer)
    txtImpostoServico(3) = txtImpostoServico(3) + cmdImpostoServico(Index).Caption + " "
    txtImpostoServico(3).SelStart = Len(txtImpostoServico(3).Text)
    txtImpostoServico(3).SetFocus
    Select Case Index
    Case 7 'INCLUIR
        If ValidaInserirCampos("FormToLV") = True Then
            If cboImpostoServico = "IMPOSTO" Then
                vPonte1.Text = 1
            Else
                vPonte1.Text = 2
            End If
            
            If Check1.Value = 1 Then
                vPonte2 = "S"
            Else
                vPonte2 = "N"
            End If
                        
            
            IncluirLV ListView1, txtImpostoServico(0), txtImpostoServico(1), vPonte1, txtImpostoServico(2), txtImpostoServico(4), txtImpostoServico(3), txtImpostoServico(5), vPonte2, txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0)
            LimpaControles txtImpostoServico(0), txtImpostoServico(1), txtImpostoServico(4), txtImpostoServico(2), txtImpostoServico(3), txtImpostoServico(5), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(3), txtImpostoServico(0)
            'cboImpostoServico.Text = ""
            txtImpostoServico(0) = Format(GeraCodigoLV(ListView1), "00")
            ConfLV
        End If
        
    Case 8 'NOVO
        LimpaControles txtImpostoServico(0), txtImpostoServico(1), txtImpostoServico(4), txtImpostoServico(2), txtImpostoServico(3), txtImpostoServico(5), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(3), txtImpostoServico(0)
        'cboImpostoServico.Text = ""
        txtImpostoServico(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 9 'EDITAR
        vPonte1.Text = cboImpostoServico.Text
        AlteraLV ListView1, txtImpostoServico(0), txtImpostoServico(1), vPonte1, txtImpostoServico(2), txtImpostoServico(4), txtImpostoServico(3), txtImpostoServico(5), vPonte2, txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0)
        If vPonte1.Text = 1 Then
            cboImpostoServico.Text = "IMPOSTO"
        Else
            cboImpostoServico.Text = "SERVIÇO"
        End If
        
        If vPonte2.Text = "S" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    Case 10 'EXCLUIR
        ExcluirItemLV ListView1
        txtImpostoServico(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 11 'SALVAR
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos e enviados com sucesso!", Ok, informacao, "ZEUS"
            'Unload Me
        Else
            SkinLabel1.Visible = False
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ZEUS"
        End If
    Case 12 'SAIR
        Unload Me
    End Select
End Sub

Private Sub Command1_Click()
    frmVariaveis.Show 1
End Sub

Private Sub Form_Load()
    inicializa_tabs SSTab1, Picture1
    exibeOpt
    
    Text1.BackColor = 12829636
    Text6.BackColor = 12829636
    listview_cabecalho
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    Compoe_ListviewVariaveis lstListView
    Compoe_ListviewMatrizes lstListView2
    
    chamaSQL "SELECT IDIMPOSTOSSERVICOS, NMIMPOSTOSSERVICOS, TIPO, DESCRICAO, F_ALIQUOTA, F_VALOR, F_VALORKG, ATIVO FROM TBIMPOSTOSSERVICOS"
    Compoe_Listview ListView1, Sqlp, "00"
    
    txtImpostoServico(0) = Format(GeraCodigoLV(ListView1), "00")
    ConfLV
    
    carregarIconBotao
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub ListView1_DblClick()
    vPonte1.Text = cboImpostoServico.Text
    AlteraLV ListView1, txtImpostoServico(0), txtImpostoServico(1), vPonte1, txtImpostoServico(2), txtImpostoServico(4), txtImpostoServico(3), txtImpostoServico(5), vPonte2, txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0), txtImpostoServico(0)
    If vPonte1.Text = 1 Then
        cboImpostoServico.Text = "IMPOSTO"
    Else
        cboImpostoServico.Text = "SERVIÇO"
    End If
    
    If vPonte2.Text = "S" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub


Private Sub Option1_Click()
    exibeOpt
End Sub

Private Sub Option2_Click()
    exibeOpt
End Sub

Private Sub Option3_Click()
    exibeOpt
End Sub

Private Sub Option4_Click()
    exibeOpt
End Sub

Private Sub txtImpostoServico_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtImpostoServico(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim x As Integer
    For x = 1 To txtImpostoServico.Count - 1
        txtImpostoServico(x).SelStart = 0
        txtImpostoServico(x).SelLength = Len(txtImpostoServico(x).Text)
    Next
End Sub

Private Sub txtImpostoServico_LostFocus(Index As Integer)
    voltaCorText txtImpostoServico(Index)
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Alíquota", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Valor", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Valor KG", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 4
    
    lstListView.ColumnHeaders.Clear
    lstListView.ColumnHeaders.Add , , "VARIÁVEIS", lstListView.Width / 1.1
    
    lstListView2.ColumnHeaders.Clear
    lstListView2.ColumnHeaders.Add , , "MATRIZES", lstListView2.Width / 1.1
    
    lstListView.View = lvwReport
    lstListView2.View = lvwReport
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Function salvar_Dados()
On Error GoTo Err
    'Grava dados ListView1
    salvar_Dados = True
    limpaQualquerDado
    desConfLV
    ordenaLVArray ListView1, "0", "1", "2", "3", "4", "5", "6", "7", "", "", "", "", "", "", "", ""
    GravaDadosLV "tbImpostosServicos", "", "I", txtImpostoServico(0)
    ConfLV
    'AtualizaListview
    Exit Function
Err:
    salvar_Dados = False
End Function

Private Function ValidaInserirCampos(FormToLV_or_LVToTable As String)
'Informe LV ou TB como parâmetro ao chamar a Function
'Para que o sistema entenda se será validado dados que serão inseridos de campos do form parav um LV: ListView ou
'Irá validar dados que serão inseridos de ListView para uma TB: Tabela do banco de dados
    If FormToLV_or_LVToTable = "FormToLV" Then
        Dim x As Integer
        ValidaInserirCampos = False
        For x = 0 To 4
            If Trim(txtImpostoServico(x).Text) = "" Then
                mobjMsg.Abrir "Favor informar o campo " & Me.txtImpostoServico(x).Tag, Ok, critico, "Atenção"
                Me.txtImpostoServico(x).SetFocus
                Exit Function
            End If
        Next
        
        If cboImpostoServico.Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.cboImpostoServico.Tag, Ok, critico, "Atenção"
            Me.cboImpostoServico.SetFocus
            Exit Function
        End If
    Else
        If ListView1.ListItems.Count = 0 Then
            mobjMsg.Abrir "Deve ser informado ao menos 01 IMPOSTO ou SERVIÇO", Ok, critico, "Atenção"
            Me.txtImpostoServico(1).SetFocus
            Exit Function
        End If
    End If
    ValidaInserirCampos = True
End Function

Private Sub lstListView_Click()
    On Error Resume Next
    Select Case vOndeEstaOTab
        Case 3, 4, 5
        AlteraLVFormulas lstListView, txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), vOndeEstaOTab
    End Select
    txtImpostoServico(vOndeEstaOTab).SelStart = Len(txtImpostoServico(vOndeEstaOTab))
    txtImpostoServico(vOndeEstaOTab).SetFocus
End Sub

Private Sub lstListView2_Click()
    On Error Resume Next
    Select Case vOndeEstaOTab
        Case 3, 4, 5
        AlteraLVFormulas lstListView2, txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), txtImpostoServico(vOndeEstaOTab), vOndeEstaOTab
    End Select
    txtImpostoServico(vOndeEstaOTab).SelStart = Len(txtImpostoServico(vOndeEstaOTab))
    txtImpostoServico(vOndeEstaOTab).SetFocus
End Sub

Private Sub txtImpostoServico_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    vOndeEstaOTab = Index
End Sub

Private Sub exibeOpt()
    If Option1.Value = True Then
        lstListView.Visible = True
        lstListView2.Visible = False
    End If
    If Option2.Value = True Then
        lstListView.Visible = False
        lstListView2.Visible = True
    End If
    If Option3.Value = True Then 'Legenda Matrizes
        Frame4.Visible = False
        Frame5.Visible = True
    End If
    If Option4.Value = True Then 'Legenda Variáveis
        Frame4.Visible = True
        Frame5.Visible = False
    End If '
End Sub

'Private Sub inicializa_tabs()
'    SSTab1.Tab = 0
'    SubClassSSTAB SSTab1, Picture1
'End Sub

Private Sub ConfLV()
    Dim x As Integer, y As Integer
    y = ListView1.ListItems.Count
    For x = 1 To y
        ListView1.ListItems(x).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(7) = "S" Then
            ListView1.SelectedItem.ListSubItems.Item(7) = ""
            ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "OK"
        ElseIf ListView1.SelectedItem.ListSubItems.Item(7) = "N" Then
            ListView1.SelectedItem.ListSubItems.Item(7) = ""
            ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "EXC"
        End If
    Next
End Sub

Private Sub desConfLV()
    Dim x As Integer, y As Integer
    y = ListView1.ListItems.Count
    For x = 1 To y
        ListView1.ListItems(x).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "OK" Then
            ListView1.SelectedItem.ListSubItems.Item(7) = "S"
        ElseIf ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "EXC" Then
            ListView1.SelectedItem.ListSubItems.Item(7) = "N"
        End If
    Next
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdImpostoServico(7), 7, 46 'Inserir
    carregaImagemBotao cmdImpostoServico(8), 8, 31 'Novo
    carregaImagemBotao cmdImpostoServico(9), 9, 32 'Editar
    carregaImagemBotao cmdImpostoServico(10), 10, 33 'Excluir
    carregaImagemBotao cmdImpostoServico(11), 11, 45 'Salvar
    carregaImagemBotao cmdImpostoServico(12), 12, 34 'Sair
End Sub
