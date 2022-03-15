VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmReceitasDespesas 
   Caption         =   "Despesas e Créditos"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10545
   Icon            =   "frmReceitasDespesas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDespesasCreditos 
      Height          =   615
      Index           =   12
      Left            =   720
      Picture         =   "frmReceitasDespesas.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   9960
      Width           =   615
   End
   Begin VB.CommandButton cmdDespesasCreditos 
      Height          =   615
      Index           =   11
      Left            =   120
      Picture         =   "frmReceitasDespesas.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "Salvar"
      ToolTipText     =   "Salvar"
      Top             =   9960
      Width           =   615
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
      TabIndex        =   33
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   31
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
         TabIndex        =   32
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
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
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdDespesasCreditos 
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
         Picture         =   "frmReceitasDespesas.frx":265E
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "Excluir"
         ToolTipText     =   "Excluir"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdDespesasCreditos 
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
         Picture         =   "frmReceitasDespesas.frx":3328
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "Editar"
         ToolTipText     =   "Editar"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdDespesasCreditos 
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
         Picture         =   "frmReceitasDespesas.frx":3FF2
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "Novo"
         ToolTipText     =   "Novo"
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox txtDespesasCreditos 
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
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "ID"
         ToolTipText     =   "Identificador do Imposto ou Serviço"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDespesasCreditos 
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
         TabIndex        =   25
         Tag             =   "Nome"
         ToolTipText     =   "Nome do imposto ou Serviço"
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtDespesasCreditos 
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
         TabIndex        =   24
         Tag             =   "Descrição"
         ToolTipText     =   "Descrição do Imposto ou Serviço"
         Top             =   1200
         Width           =   10095
      End
      Begin VB.CommandButton cmdDespesasCreditos 
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
         Picture         =   "frmReceitasDespesas.frx":4CBC
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "Incluir"
         ToolTipText     =   "Incluir"
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
         TabIndex        =   21
         ToolTipText     =   "Seleciono o Tipo: IMPOSTO ou SERVIÇO"
         Top             =   240
         Width           =   2055
         Begin VB.ComboBox cboDespesasCreditos 
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
            ItemData        =   "frmReceitasDespesas.frx":5986
            Left            =   120
            List            =   "frmReceitasDespesas.frx":5990
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Tipo: IMPOSTO ou SERVIÇO"
            ToolTipText     =   "Selecione o Tipo: IMPOSTO ou SERVIÇO"
            Top             =   360
            Width           =   1815
         End
      End
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
               Picture         =   "frmReceitasDespesas.frx":59A6
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReceitasDespesas.frx":63B8
               Key             =   "EXC"
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
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
         TabPicture(0)   =   "frmReceitasDespesas.frx":6DCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Legenda"
         TabPicture(1)   =   "frmReceitasDespesas.frx":6DE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame7"
         Tab(1).ControlCount=   1
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
            TabIndex        =   9
            Top             =   360
            Width           =   9855
            Begin VB.TextBox txtDespesasCreditos 
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
               TabIndex        =   13
               Tag             =   "Alíquota"
               ToolTipText     =   "Percentual a ser Aplicado sobre o valor do orçamento"
               Top             =   480
               Width           =   5415
            End
            Begin VB.TextBox txtDespesasCreditos 
               Height          =   345
               Index           =   3
               Left            =   4320
               TabIndex        =   12
               Tag             =   "Compor Fórmula"
               ToolTipText     =   "Composição da Fórmula do IMPOSTO ou SERVIÇO"
               Top             =   1320
               Width           =   5415
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
               TabIndex        =   11
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
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
               TabIndex        =   10
               Top             =   360
               Width           =   1695
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "frmReceitasDespesas.frx":6E02
               TabIndex        =   14
               Top             =   240
               Width           =   3015
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "frmReceitasDespesas.frx":6E72
               TabIndex        =   15
               Top             =   1080
               Width           =   1815
            End
            Begin MSComctlLib.ListView lstListView 
               Height          =   2655
               Left            =   120
               TabIndex        =   16
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
               TabIndex        =   17
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
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3495
            Left            =   -74880
            TabIndex        =   2
            Top             =   360
            Width           =   9855
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
               TabIndex        =   8
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
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
               TabIndex        =   7
               Top             =   120
               Width           =   1695
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
               TabIndex        =   5
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
                  TabIndex        =   6
                  Text            =   "frmReceitasDespesas.frx":6EEE
                  ToolTipText     =   "Legenda referente aos objetos que podem ser utilizados na fórmula"
                  Top             =   240
                  Width           =   9375
               End
            End
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
               TabIndex        =   3
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
                  TabIndex        =   4
                  Text            =   "frmReceitasDespesas.frx":72F5
                  ToolTipText     =   "Legenda referente aos objetos que podem ser utilizados na fórmula"
                  Top             =   240
                  Width           =   9375
               End
            End
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReceitasDespesas.frx":7E97
         TabIndex        =   27
         Top             =   960
         Width           =   3735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmReceitasDespesas.frx":7F39
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmReceitasDespesas.frx":7F9B
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   30
         Tag             =   "Itens do Sub-critério"
         ToolTipText     =   "Lista de IMPOSTOS e SERVIÇOS cadastrados"
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
Attribute VB_Name = "frmReceitasDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DESPESAS E CREDITOS
Private vPonte1 As TextBox, vPonte2 As TextBox
Private rsCriterio As New ADODB.Recordset
Private SqlCriterio As String
Private vOndeEstaOTab As Integer

Private Sub cmdDespesasCreditos_Click(Index As Integer)
    txtDespesasCreditos(3) = txtDespesasCreditos(3) + cmdDespesasCreditos(Index).Caption + " "
    txtDespesasCreditos(3).SelStart = Len(txtDespesasCreditos(3).Text)
    txtDespesasCreditos(3).SetFocus
    Select Case Index
    Case 7 'INCLUIR
        If ValidaInserirCampos("FormToLV") = True Then
            If cboDespesasCreditos = "DESPESA" Then
                vPonte1.Text = 1
            Else
                vPonte1.Text = 2
            End If
            
            If Check1.Value = 1 Then
                vPonte2 = "S"
            Else
                vPonte2 = "N"
            End If
                        
            
            IncluirLV ListView1, txtDespesasCreditos(0), txtDespesasCreditos(1), vPonte1, txtDespesasCreditos(2), txtDespesasCreditos(4), txtDespesasCreditos(3), vPonte2, txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0)
            LimpaControles txtDespesasCreditos(0), txtDespesasCreditos(1), txtDespesasCreditos(4), txtDespesasCreditos(2), txtDespesasCreditos(3), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(3), txtDespesasCreditos(0), txtDespesasCreditos(0)
            'cboDespesasCreditos.Text = ""
            txtDespesasCreditos(0) = Format(GeraCodigoLV(ListView1), "00")
            ConfLV
        End If
        
    Case 8 'NOVO
        LimpaControles txtDespesasCreditos(0), txtDespesasCreditos(1), txtDespesasCreditos(4), txtDespesasCreditos(2), txtDespesasCreditos(3), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(3), txtDespesasCreditos(0), txtDespesasCreditos(0)
        'cboDespesasCreditos.Text = ""
        txtDespesasCreditos(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 9 'EDITAR
        vPonte1.Text = cboDespesasCreditos.Text
        AlteraLV ListView1, txtDespesasCreditos(0), txtDespesasCreditos(1), vPonte1, txtDespesasCreditos(2), txtDespesasCreditos(4), txtDespesasCreditos(3), vPonte2, txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0)
        If vPonte1.Text = 1 Then
            cboDespesasCreditos.Text = "DESPESA"
        Else
            cboDespesasCreditos.Text = "CRÉDITO"
        End If
        
        If vPonte2.Text = "S" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    Case 10 'EXCLUIR
        ExcluirItemLV ListView1
        txtDespesasCreditos(0) = Format(GeraCodigoLV(ListView1), "00")
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
    inicializa_tabs
    exibeOpt
    
    Text1.BackColor = 12829636
    Text6.BackColor = 12829636
    listview_cabecalho
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    Compoe_ListviewVariaveis lstListView
    Compoe_ListviewMatrizes lstListView2
    
    chamaSQL "SELECT IDDESPESASCREDITOS, CONTA, TIPO, DESCRICAO, F_ALIQUOTA, F_TOTAL, ATIVO FROM TBDESPESASCREDITOS"
    Compoe_Listview ListView1, Sqlp, "00"
    
    txtDespesasCreditos(0) = Format(GeraCodigoLV(ListView1), "00")
    ConfLV
    carregarIconBotao
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdDespesasCreditos(7), 7, 46 'Inserir
    carregaImagemBotao cmdDespesasCreditos(8), 8, 31 'Novo
    carregaImagemBotao cmdDespesasCreditos(9), 9, 32 'Editar
    carregaImagemBotao cmdDespesasCreditos(10), 10, 33 'Excluir
    carregaImagemBotao cmdDespesasCreditos(11), 11, 45 'Salvar
    carregaImagemBotao cmdDespesasCreditos(12), 12, 34 'Sair
End Sub

Private Sub ListView1_DblClick()
    vPonte1.Text = cboDespesasCreditos.Text
    AlteraLV ListView1, txtDespesasCreditos(0), txtDespesasCreditos(1), vPonte1, txtDespesasCreditos(2), txtDespesasCreditos(4), txtDespesasCreditos(3), vPonte2, txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0), txtDespesasCreditos(0)
    If vPonte1.Text = 1 Then
        cboDespesasCreditos.Text = "DESPESA"
    Else
        cboDespesasCreditos.Text = "CRÉDITO"
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

Private Sub txtDespesasCreditos_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtDespesasCreditos(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtDespesasCreditos.Count - 1
        txtDespesasCreditos(X).SelStart = 0
        txtDespesasCreditos(X).SelLength = Len(txtDespesasCreditos(X).Text)
    Next
End Sub

Private Sub txtDespesasCreditos_LostFocus(Index As Integer)
    voltaCorText txtDespesasCreditos(Index)
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Conta", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Alíquota", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Total", ListView1.Width / 4
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
    ordenaLVArray ListView1, "0", "1", "2", "3", "4", "5", "6", "", "", "", "", "", "", "", "", ""
    GravaDadosLV "tbDespesasCreditos", "", "I", txtDespesasCreditos(0)
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
        Dim X As Integer
        ValidaInserirCampos = False
        For X = 0 To 4
            If Trim(txtDespesasCreditos(X).Text) = "" Then
                mobjMsg.Abrir "Favor informar o campo " & Me.txtDespesasCreditos(X).Tag, Ok, critico, "Atenção"
                Me.txtDespesasCreditos(X).SetFocus
                Exit Function
            End If
        Next
        
        If cboDespesasCreditos.Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.cboDespesasCreditos.Tag, Ok, critico, "Atenção"
            Me.cboDespesasCreditos.SetFocus
            Exit Function
        End If
    Else
        If ListView1.ListItems.Count = 0 Then
            mobjMsg.Abrir "Deve ser informado ao menos 01 DESPESA ou CRÉDITO", Ok, critico, "Atenção"
            Me.txtDespesasCreditos(1).SetFocus
            Exit Function
        End If
    End If
    ValidaInserirCampos = True
End Function

Private Sub lstListView_Click()
    On Error Resume Next
    Select Case vOndeEstaOTab
        Case 3, 4, 5
        AlteraLVFormulas lstListView, txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), vOndeEstaOTab
    End Select
    txtDespesasCreditos(vOndeEstaOTab).SelStart = Len(txtDespesasCreditos(vOndeEstaOTab))
    txtDespesasCreditos(vOndeEstaOTab).SetFocus
End Sub

Private Sub lstListView2_Click()
    On Error Resume Next
    Select Case vOndeEstaOTab
        Case 3, 4, 5
        AlteraLVFormulas lstListView2, txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), txtDespesasCreditos(vOndeEstaOTab), vOndeEstaOTab
    End Select
    txtDespesasCreditos(vOndeEstaOTab).SelStart = Len(txtDespesasCreditos(vOndeEstaOTab))
    txtDespesasCreditos(vOndeEstaOTab).SetFocus
End Sub

Private Sub txtDespesasCreditos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub inicializa_tabs()
    SSTab1.Tab = 0
    SubClassSSTAB SSTab1, Picture1
End Sub

Private Sub ConfLV()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems(X).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(6) = "S" Then
            ListView1.SelectedItem.ListSubItems.Item(6) = ""
            ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK"
        ElseIf ListView1.SelectedItem.ListSubItems.Item(6) = "N" Then
            ListView1.SelectedItem.ListSubItems.Item(6) = ""
            ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "EXC"
        End If
    Next
End Sub

Private Sub desConfLV()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems(X).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK" Then
            ListView1.SelectedItem.ListSubItems.Item(6) = "S"
        ElseIf ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "EXC" Then
            ListView1.SelectedItem.ListSubItems.Item(6) = "N"
        End If
    Next
End Sub

