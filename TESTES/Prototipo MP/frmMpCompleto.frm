VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Begin VB.Form frmMPCompleto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metodos e Processos"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16065
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMpCompleto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   16065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox label53 
      Height          =   330
      Left            =   7080
      TabIndex        =   67
      Text            =   "-"
      Top             =   9600
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sequencial "
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
      Left            =   14640
      TabIndex        =   65
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtformula 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         HideSelection   =   0   'False
         Index           =   15
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "ID"
         Top             =   360
         Width           =   1095
      End
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
      Index           =   13
      Left            =   720
      Picture         =   "frmMpCompleto.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9480
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
      Index           =   12
      Left            =   120
      Picture         =   "frmMpCompleto.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9480
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Desenhos"
      TabPicture(0)   =   "frmMpCompleto.frx":265E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame31"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCadastro(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCadastro(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCadastro(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Recursos"
      TabPicture(1)   =   "frmMpCompleto.frx":267A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame9"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ListView1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdCadastro(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtformula(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtformula(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdCadastro(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdCadastro(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdCadastro(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame6"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Frame7"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "ScriptControl1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Frame13"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Frame14"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      Begin VB.Frame Frame14 
         Caption         =   "Tempo total "
         Height          =   615
         Left            =   9840
         TabIndex        =   68
         Top             =   7560
         Width           =   4215
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   195
            Width           =   2895
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Grupo "
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
         TabIndex        =   62
         Top             =   4080
         Width           =   4215
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Tag             =   "Grupo"
            ToolTipText     =   "Grupo"
            Top             =   240
            Width           =   3975
         End
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   10560
         Top             =   4200
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.Frame Frame7 
         Caption         =   "Data Prevista"
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
         Left            =   4800
         TabIndex        =   48
         Top             =   3240
         Width           =   1815
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   405
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            _Version        =   393216
            Format          =   16252929
            CurrentDate     =   41556
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tempo calculado (Min)"
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
         Left            =   2400
         TabIndex        =   46
         Top             =   3240
         Width           =   2295
         Begin VB.TextBox txtResultado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   525
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   47
            Tag             =   "Tempo calculado"
            ToolTipText     =   "Tempo calculado"
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   4
         Left            =   -67440
         Picture         =   "frmMpCompleto.frx":2696
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "Limpar Controles"
         ToolTipText     =   "Limpar Controles"
         Top             =   7200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1320
         Picture         =   "frmMpCompleto.frx":3360
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   720
         Picture         =   "frmMpCompleto.frx":402A
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtformula 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "ID Centro de Custo"
         ToolTipText     =   "ID Centro de Custo"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtformula 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2760
         TabIndex        =   15
         Top             =   600
         Width           =   7575
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmMpCompleto.frx":4CF4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4200
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Figura "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   6720
         TabIndex        =   39
         Top             =   960
         Width           =   4455
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   3495
            Left            =   120
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   6165
            Image           =   "frmMpCompleto.frx":59BE
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fórmulas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   11280
         TabIndex        =   37
         Top             =   360
         Width           =   4455
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   38
            Top             =   3960
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSComctlLib.TreeView TreeView3 
            Height          =   4095
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   7223
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Observação "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   6495
         Begin VB.TextBox txtformula 
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H000000C0&
            Height          =   1815
            Index           =   6
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   36
            Tag             =   "Fórmula"
            ToolTipText     =   "Fórmula"
            Top             =   240
            Width           =   6255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Variáveis "
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
         Left            =   120
         TabIndex        =   34
         Top             =   3240
         Width           =   2175
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Tag             =   "Insira as variáveis de acordo com a Observação acima"
            ToolTipText     =   "Insira as variáveis de acordo com a Observação acima"
            Top             =   360
            Width           =   1935
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   20
         Top             =   4920
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   4683
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "<"
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
         Index           =   8
         Left            =   -67560
         TabIndex        =   11
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   ">"
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
         Index           =   7
         Left            =   -67560
         TabIndex        =   9
         Top             =   2400
         Width           =   735
      End
      Begin VB.Frame Frame31 
         Caption         =   "Itens selecionados "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -66720
         TabIndex        =   31
         Top             =   480
         Width           =   7335
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   6240
            Visible         =   0   'False
            Width           =   7095
         End
         Begin VB.Frame Frame3 
            Caption         =   "Peso Total Selecionado"
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
            TabIndex        =   44
            Top             =   6600
            Width           =   7095
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   6855
            End
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   6255
            Left            =   120
            TabIndex        =   10
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   11033
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
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
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Desenhos disponíveis "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   7215
         Begin VB.Frame Frame41 
            Caption         =   "Peso Total Selecionado"
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
            TabIndex        =   32
            Top             =   6600
            Width           =   6975
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   6735
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   6255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   11033
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
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
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Referências "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   10
            Left            =   2160
            TabIndex        =   55
            Top             =   600
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   7
            Left            =   2160
            TabIndex        =   52
            Top             =   600
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   50
            Top             =   480
            Width           =   5415
         End
         Begin VB.Label Label7 
            Caption         =   "Parâmetros:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Fórmula:"
            Height          =   255
            Left            =   2160
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Decoder "
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
         TabIndex        =   58
         Top             =   1920
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtDecoder 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Constantes "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   11280
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin MSComctlLib.ListView ListView2 
            Height          =   2895
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   5106
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   4210752
            BackColor       =   16777215
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
      Begin VB.Label Label1 
         Caption         =   "ID C.C.:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nome C.C.:"
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame21 
      Caption         =   "Dados Método e Processo "
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
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   14415
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   13800
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   14
         Left            =   9960
         TabIndex        =   6
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   9240
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   13
         Left            =   5040
         TabIndex        =   4
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtformula 
         Height          =   285
         Index           =   12
         Left            =   3240
         TabIndex        =   2
         Text            =   "4233"
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Format          =   16252929
         CurrentDate     =   41554
      End
      Begin VB.TextBox txtformula 
         Enabled         =   0   'False
         Height          =   285
         HelpContextID   =   1
         Index           =   11
         Left            =   120
         TabIndex        =   0
         Text            =   "000001"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label51 
         Caption         =   "Responsável"
         Height          =   255
         Left            =   9960
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Projeto:"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label31 
         Caption         =   "FCE:"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Data:"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Programação nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMPCompleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variaveis que irao receber os valores referente aos parametros das formulas
'para localizar os dados na tabela de classificação

'Variaveis que irão receber os dados da tabela de classificação após a localizacao
Private vTMedio As Double '
Private vFFadiga As Double
Private vOrganiza As Double
Private vSomaTempo As Double

'Variáveis que irão receber os dados do textBox de parametro para realizar a localização na
'tabela de parametros
Private vGrupo As String
Private vDimTipo As String
Private vDimValor As String
Private vInterTipo As String
Private vInterValor As String

Private var(50) As Double
Private cons(50) As Double
'---------------------------------------------------

Private vNomeA As String
Private vNomeB As String
Private vNomeC As String
Private vJuntaNome As String
Private vPesoTotal1 As Double
Private vPesoTotal2 As Double
Private vAcumula As String
Private vNmNo As String

Private vPonte1 As TextBox
Private vPonte2 As TextBox
Private vPonte3 As TextBox

Private Sub cmdCadastro_Click(Index As Integer)
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    Select Case Index
    Case 0
        'Chama CC - Centro de Custo
        Label8 = "-"
        ChamaGrid "tbCCusto", "nome", txtformula(0), frmFormulaCC, "idprd", "nome"
        CarregaTxt "tbCCusto", "idprd", "S", "", "", txtformula(0), txtformula(1), 0, 1, txtformula(0), "S", txtformula(1)
        montaEstrutTreeview
        
        LimpaVariaveis
        LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
        
        compoeDadosLVs
    Case 1
        'Cria TextBox em tempo de Execução
        
        vPonte1.Text = "-"
        vPonte2.Text = DTPicker2.Value
        vPonte3.Text = Label8.Caption
        
        If ValidaCampos(ListView1, txtformula(0), txtformula(6), Text1, vPonte3, txtformula(5), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0)) = False Then Exit Sub
        IncluirLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5)
        
        LimpaVariaveis
        LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
        txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        SomaLV ListView1
    Case 2
        AlteraLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5)
        montaEstrutTreeview
        compoeDadosLVs
        DTPicker2.Value = vPonte2.Text
        Label8.Caption = vPonte3.Text
        EditaTreeview
        compoeControles
        separaDadosText1 Text1
        Text1 = ""
        mostraDesenhos "tbMPDesSel", TreeView2
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 3
        ExcluirItemLV ListView1
        LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
        txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        SomaLV ListView1
    Case 4
        TreeView2.Nodes.Clear
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        Label3 = "-"
    Case 7
        sqlDeletar = "Delete from tbMPDesSel"
        rsDeletar.Open sqlDeletar, cnBanco
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        buscaChecado2 TreeView1
        mostraDesenhos "tbMPDesSel", TreeView2
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 8
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        buscaChecado2 TreeView2
        mostraDesenhos "tbMPDesSel", TreeView2
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 13 'Sair do formulário
        Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    Set vPonte3 = Me.Controls.Add("VB.TextBox", "vPonte3")
    DTPicker1 = Date
    DTPicker2 = Date
    SSTab1.Tab = 0
    listview_cabecalho
    mostraDesenhos "tbitemlm", TreeView1
    txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
End Sub

Private Sub mostraDesenhos(vTabela As String, TV As TreeView)
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim vNome1 As String, vNome2 As String, vNome3 As String
    Dim nd As Node
    Dim vPula As Integer
    Dim vNo As Integer, vNo2 As Integer
    Dim vNomeNo As String
       
    TV.Nodes.Clear

    If vTabela = "tbitemlm" Then
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq " & _
        "from tbitemlm as a inner join CORPORERM.dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join CORPORERM.dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT Where a.fce = '" & Val(txtformula(12)) & "' Order by c.desenho,d.posicao,d.item"
    ElseIf vTabela = "tbMPDesSel" Then
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq " & _
        "from tbitemlm as a inner join CORPORERM.dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join CORPORERM.dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbMPDesSel as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq Where a.fce = '" & Val(txtformula(12)) & "' Order by c.desenho,d.posicao,d.item"
    End If
    
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
    vJuntaNome = ""
    vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & ";" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
    separaDadosTree vJuntaNome
    vNome1 = vNomeA
    vNome2 = vNomeB
    vNome3 = vNomeC
    vNo = 0
    On Error Resume Next
    Do While Not rsTreeview.EOF
        'PRIMEIRO NO
        Set nd = TV.Nodes.Add(, , vNome1, vNome1)
        Do While vNome1 = vNomeA And Not rsTreeview.EOF
            If vNomeB <> "" Then
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
                'SEGUNDO NO
                Set nd = TV.Nodes.Add(vNome1, tvwChild, vNomeNo, vNome2)
                Do While vNome2 = vNomeB And vNomeC <> "" And Not rsTreeview.EOF
                    'TERCEIRO NO
                    'OBS: OS VALORES DOS NOs NÃO PODEM SE REPETIR
                    'FOI ADICIONADO UM CONTADOR AO IDENTIFICADOR DO NO PARA QUE ELE NÃO SE REPITA
                    If TV.Name = "TreeView2" Then
                        vPesoTotal2 = vPesoTotal2 + (rsTreeview.Fields(6) * rsTreeview.Fields(7) * rsTreeview.Fields(9))
                        If Text1.Text = "" Then
                            Text1 = Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                        Else
                            Text1 = Text1.Text & ";" & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                        End If
                    End If
                    Set nd = TV.Nodes.Add(vNomeNo, tvwChild, vNomeC & vNo, vNomeC)
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                    vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & ";" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                    separaDadosTree vJuntaNome
                    vPula = 1
                Loop
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
            End If
            If vPula = 0 Then
                If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & ";" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                separaDadosTree vJuntaNome
            End If
            vPula = 0
            
            If Not rsTreeview.EOF Then
                vNome2 = vNomeB
            End If
        Loop
        If Not rsTreeview.EOF Then
            vNome1 = vNomeA
            vNome2 = vNomeB
            vNome3 = vNomeC
        End If
    Loop
    rsTreeview.Close
    Set rsTreeview = Nothing
End Sub

Private Sub ListView1_DblClick()
    AlteraLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5)
    montaEstrutTreeview
    compoeDadosLVs
    DTPicker2.Value = vPonte2.Text
    Label8.Caption = vPonte3.Text
    EditaTreeview
    compoeControles
    separaDadosText1 Text1
    Text1 = ""
    mostraDesenhos "tbMPDesSel", TreeView2
    If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim aux As MSComctlLib.Node
    Set aux = Node.Child
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        If Not aux.Child Is Nothing Then
            TreeView1_NodeCheck aux
        End If
        Set aux = aux.Next
    Loop
    Set aux = Node.Parent
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        Set aux = aux.Parent
    Loop
    
    vAcumula = ""
    Label6 = "-"
    vPesoTotal1 = 0
    buscaChecado
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim aux As MSComctlLib.Node
    Set aux = Node.Child
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        If Not aux.Child Is Nothing Then
            TreeView2_NodeCheck aux
        End If
        Set aux = aux.Next
    Loop
    Set aux = Node.Parent
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        Set aux = aux.Parent
    Loop
    'vPesoTotal = 0
End Sub

Private Sub buscaChecado()
    Dim X As Integer, vContador As Integer, vQtdNos As Integer
    vContador = 0
    X = 0
    vQtdNos = TreeView1.Nodes.Count
    For X = 1 To vQtdNos
        If TreeView1.Nodes.Item(X).Checked = True Then
            PegaTreeview X
            separaDadosTree vJuntaNome
            buscaPeso
        End If
    Next
End Sub

Private Sub buscaChecado2(vLV As TreeView)
    Dim X As Integer, vContador As Integer, vQtdNos As Integer
    vContador = 0
    X = 0
    vQtdNos = vLV.Nodes.Count
    For X = 1 To vQtdNos
        If vLV.Nodes.Item(X).Checked = True Then
            transfDesenhosSel X, vLV
        End If
    Next
End Sub

Private Sub PegaTreeview(llng_Contador As Integer)
    If TreeView1.Nodes(llng_Contador).Checked = True Then
        vNmNo = TreeView1.Nodes(llng_Contador).FullPath
    End If
    vNmNo = Replace(vNmNo, "\", ";")
    vJuntaNome = vNmNo
End Sub

Private Sub buscaPeso()
    Dim rsBuscaPeso As New ADODB.Recordset
    Dim SqlBuscaPeso As String
    Dim vCodLM As String, vCodSeq As String
        
    If vNomeC <> "" Then
        vNomeC = Right(vNomeC, 5)
        vCodLM = Mid$(vNomeC, 1, 2)
        vCodSeq = Mid$(vNomeC, 3, 3)
        
        If vAcumula = vNomeC And Label6 <> "-" Then
            Exit Sub
        Else
            vAcumula = vNomeC
        End If
        
        SqlBuscaPeso = "select a.quantcj*a.quantunit*a.pesounit as PesoTotal from tbItemLM as a where a.fce = '" & Val(txtformula(12)) & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "'"
        rsBuscaPeso.Open SqlBuscaPeso, cnBanco, adOpenKeyset, adLockReadOnly
        vPesoTotal1 = vPesoTotal1 + rsBuscaPeso.Fields(0)
    End If
    Label6 = Format(vPesoTotal1, "#,##0.00;(#,##0.00)")
End Sub

'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Seq.", ListView2.Width / 5
    ListView1.ColumnHeaders.Add , , "OP nº", ListView2.Width / 6
    ListView1.ColumnHeaders.Add , , "ID. C.Custo", ListView2.Width / 2.5
    ListView1.ColumnHeaders.Add , , "Nome C. Custo", ListView2.Width / 1.2
    ListView1.ColumnHeaders.Add , , "Desenhos/Items", ListView2.Width / 1
    ListView1.ColumnHeaders.Add , , "Data Prevista", ListView2.Width / 2.5
    ListView1.ColumnHeaders.Add , , "T. Calculado", ListView2.Width / 3.5
    ListView1.ColumnHeaders.Add , , "Grupo", ListView2.Width / 10000
    ListView1.ColumnHeaders.Add , , "ID Programação", ListView2.Width / 10000
    ListView1.ColumnHeaders.Add , , "Variáveis", ListView2.Width / 3.5
    
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Valor constante", ListView2.Width / 2.5
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 3.5
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    
End Sub

Private Sub compoeControles()
    Dim rsCompoe As New ADODB.Recordset
    Dim SqlCompoe As String
    'SqlCompoe = "Select a.parametros,a.formula from tbFormula as a where a.idprd = '" & txtformula(0) & "' and idform = '" & Val(txtformula(4)) & "'"
    SqlCompoe = "Select a.parametros,a.formula,a.observacao,a.imagem from tbFormula as a inner join tbproduto as b on a.idprd = b.idprd where a.idprd = '" & txtformula(0) & "' and idform = '" & Val(txtformula(4)) & "'"
    rsCompoe.Open SqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
        txtformula(2).Text = rsCompoe.Fields(0) 'Parâmetros
        txtformula(3).Text = rsCompoe.Fields(1) 'Formula
        If Not IsNull(rsCompoe.Fields(2)) Then txtformula(6).Text = rsCompoe.Fields(2) 'Observação
        If Not IsNull(rsCompoe.Fields(3)) Then label53 = rsCompoe.Fields(3) Else label53 = "" 'Imagem
    Else
        txtformula(2).Text = "" 'Parâmetros
        txtformula(3).Text = "" 'Formula
        txtformula(6).Text = "" 'Observação
        label53 = "" 'Imagem
    End If
    If Mid$(txtformula(2).Text, 1, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 9, 1), 1
    End If
    If Mid$(txtformula(2).Text, 12, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 20, 1), 2
    End If
    
    separaDadosTree vNmNo
    If vNomeC <> "" Then
        Label8 = vNomeA & "/" & vNomeB & "/" & vNomeC
    ElseIf vNomeC = "" And vNomeB <> "" Then
        Label8 = vNomeA & "/" & vNomeB
    ElseIf vNomeB = "" Then
        Label8 = vNomeA
    End If
    
    aicAlphaImage1.ClearImage
    If label53 <> "" Or label53 <> "-" Then
        aicAlphaImage1.LoadImage_FromFile (label53.Text)
    End If
    
    rsCompoe.Close
    Set rsCompoe = Nothing
End Sub

Private Sub compoeDadosLVs()
    'Faz referências a Funções que estão no: Module1.bas
    'Listview2 - Constantes
    LimpaLV ListView2
    chamaSQL "Select a.idseq,a.valconst,a.descricao from tbconstantes as a where a.idprd = '" & txtformula(0) & "'"
    Compoe_Listview ListView2, Sqlp, "000"
End Sub

Private Sub LimpaVariaveis()
    vGrupo = ""
    vDimTipo = ""
    vDimValor = ""
    vInterTipo = ""
    vInterValor = ""
    vSomaTempo = 0
    vTMedio = 0
    vFFadiga = 0
    vOrganiza = 0
    vSomaTempo = 0
End Sub

'As 3 próximas SUBs são referentes a montagem e manipulação do TREEVIEW3
Private Sub montaEstrutTreeview()
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim vNome1 As String, vNome2 As String, vNome3 As String
    Dim nd As Node
    Dim vPula As Integer
    Dim vNo As Integer, vNo2 As Integer
    Dim vNomeNo As String
       
    TreeView3.Nodes.Clear

    SqlTreeview = "Select * from tbFormula as a where a.idprd = '" & txtformula(0) & "' order by a.idprd,a.nmform"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
    separaDadosTree rsTreeview.Fields(2)
    vNome1 = vNomeA
    vNome2 = vNomeB
    vNome3 = vNomeC
    vNo = 0
    On Error Resume Next
    Do While Not rsTreeview.EOF
        'PRIMEIRO NO
        Set nd = TreeView3.Nodes.Add(, , vNome1, vNome1)
        Do While vNome1 = vNomeA And Not rsTreeview.EOF
            If vNomeB <> "" Then
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
                'SEGUNDO NO
                Set nd = TreeView3.Nodes.Add(vNome1, tvwChild, vNomeNo, vNome2)
                Do While vNome2 = vNomeB And vNomeC <> "" And Not rsTreeview.EOF
                    'TERCEIRO NO
                    'OBS: OS VALORES DOS NOs NÃO PODEM SE REPETIR
                    'FOI ADICIONADO UM CONTADOR AO IDENTIFICADOR DO NO PARA QUE ELE NÃO SE REPITA
                    Set nd = TreeView3.Nodes.Add(vNomeNo, tvwChild, vNomeC & vNo, vNomeC)
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                    separaDadosTree rsTreeview.Fields(2)
                    vPula = 1
                Loop
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
            End If
            If vPula = 0 Then
                If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                separaDadosTree rsTreeview.Fields(2)
            End If
            vPula = 0
            
            If Not rsTreeview.EOF Then
                vNome2 = vNomeB
            End If
        Loop
        If Not rsTreeview.EOF Then
            vNome1 = vNomeA
            vNome2 = vNomeB
            vNome3 = vNomeC
        End If
    Loop
    rsTreeview.Close
    Set rsTreeview = Nothing
End Sub

Private Sub separaDadosTree(vTxtForm As String)
    Dim RECEBE As String
    Dim CONTADOR As Integer, X As Integer
    CONTADOR = 0
    vNomeA = ""
    vNomeB = ""
    vNomeC = ""
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If CONTADOR = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If CONTADOR = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
End Sub

'A função abaixo pega os valores dos parâmetro informados no textBox e armazena em variáveis
'específicas para cada valor
Private Sub separaDadosPar(vTxtForm As TextBox)
    Dim RECEBE As String
    Dim CONTADOR As Integer, vNum As Integer
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If CONTADOR = 0 And RECEBE <> "-" Then vGrupo = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 1 Then vDimTipo = RECEBE 'Variável vDimTipo receber o valor do segundo parâmetro
            If CONTADOR = 2 Then vDimValor = RECEBE 'Variavel vDimTipo recebe o valor do terceiro parâmetro
            If CONTADOR = 3 Then vInterTipo = RECEBE 'Variável vInterTipo recebe o valor do quarto parâmetro
            If CONTADOR = 4 Then vInterValor = RECEBE 'Variável vInterValor recebe o valor do quinto parâmetro
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If CONTADOR = 0 And RECEBE <> "-" Then vGrupo = RECEBE
    If CONTADOR = 1 Then vDimTipo = RECEBE
    If CONTADOR = 2 Then vDimValor = RECEBE
    If CONTADOR = 3 Then vInterTipo = RECEBE
    If CONTADOR = 4 Then vInterValor = RECEBE
    
    If Mid$(vDimValor, 1, 3) = "var" Then
        vNum = Val(Mid$(vDimValor, 5, 2))
        vDimValor = var(Val(Mid$(vDimValor, 5, 2)))
        vDimValor = Replace(vDimValor, ",", ".")
    End If
    If Mid$(vInterValor, 1, 3) = "var" Then
        vNum = Val(Mid$(vDimValor, 5, 2))
        vInterValor = var(Val(Mid$(vInterValor, 5, 2)))
        vInterValor = Replace(vInterValor, ",", ".")
    End If
End Sub

'A função abaixo pega os valores das variáveis informados no textBox txtformula(5) e armazena em Arrays: var(?)
'específicas para cada valor
Private Sub separaDadosVar(vTxtForm As TextBox)
    Dim RECEBE As String
    Dim CONTADOR As Integer, X As Integer
    CONTADOR = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If CONTADOR = 0 Then var(1) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 1 Then var(2) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 2 Then var(3) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 3 Then var(4) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 4 Then var(5) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If CONTADOR = 5 Then var(6) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If CONTADOR = 0 Then var(1) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 1 Then var(2) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 2 Then var(3) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 3 Then var(4) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 4 Then var(5) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If CONTADOR = 5 Then var(6) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
End Sub

'A função abaixo pega os valores das constantes informados no Listview2 e armazena em Arrays: cons(?)
'específicas para cada valor
Private Sub separaDadosCons()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True
        If ListView2.ListItems.Item(X).Selected = True Then
            cons(Val(ListView2.ListItems.Item(X))) = ListView2.SelectedItem.ListSubItems.Item(1)
        End If
    Next
End Sub

'A função abaixo separa os valores do texbox TEXT1 e grava na tabela tbMPDesSel
Private Sub separaDadosText1(vTxtForm As TextBox)
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    Dim vCodLM As String, vCodSeq As String
    
    SqlTransf = "Delete from tbMPDesSel where fce = '" & Val(txtformula(12)) & "'"
    rsTransf.Open SqlTransf, cnBanco
    
    Dim RECEBE As String
    Dim CONTADOR As Integer, X As Integer
    CONTADOR = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 3)
            SqlTransf = "Insert into tbMPDesSel(fce,codlm,codseq) Values('" & Val(txtformula(12)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
            rsTransf.Open SqlTransf, cnBanco
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        vCodLM = Mid$(RECEBE, 1, 2)
        vCodSeq = Mid$(RECEBE, 3, 3)
        SqlTransf = "Insert into tbMPDesSel(fce,codlm,codseq) Values('" & Val(txtformula(12)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
        rsTransf.Open SqlTransf, cnBanco
    End If
End Sub

'Localiza a classificação na tabela baseado nos dados capturados na função separaDados
Private Sub localizaClassificacao()
    Dim rsLocaliza As New ADODB.Recordset
    Dim SqlLocaliza As String
    If vInterValor <> "" Then
        SqlLocaliza = "select * from tbClassificacao where idprd = '" & txtformula(0) & "' and idgrupo = '" & Val(vGrupo) & "' and '" & vDimValor & "' BETWEEN dim1 and dim2 AND '" & vInterValor & "' BETWEEN inter1 and inter2"
    End If
    If vInterValor = "" And vDimValor <> "" Then
        SqlLocaliza = "select * from tbClassificacao where idprd = '" & txtformula(0) & "' and idgrupo = '" & Val(vGrupo) & "' and '" & vDimValor & "' BETWEEN dim1 and dim2"
    End If
    
    If SqlLocaliza <> "" Then
        rsLocaliza.Open SqlLocaliza, cnBanco, adOpenKeyset, adLockReadOnly
        If Not rsLocaliza.EOF Then
            vTMedio = rsLocaliza.Fields(7)
            vFFadiga = rsLocaliza.Fields(8)
            vOrganiza = rsLocaliza.Fields(9)
            vSomaTempo = vSomaTempo + (var(2) / vTMedio)
            rsLocaliza.Close
            Set rsLocaliza = Nothing
        End If
    End If
End Sub

Private Sub AlteraTreeview()
    Dim rsAlteraTreeview As New ADODB.Recordset
    Dim SqlAlteraTreeview As String
    
    Dim llng_Contador As Long
    
    For llng_Contador = 1 To TreeView3.Nodes.Count
        If TreeView3.Nodes(llng_Contador).Selected = True Then
            vNmNo = TreeView3.Nodes(llng_Contador).FullPath
        End If
    Next
    vNmNo = Replace(vNmNo, "\", ";")
    
    SqlAlteraTreeview = "Select idform,nmform,formula,parametros from tbFormula where nmform = '" & vNmNo & "'"
    rsAlteraTreeview.Open SqlAlteraTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsAlteraTreeview.EOF Then txtformula(4) = rsAlteraTreeview.Fields(0)
End Sub

Private Sub EditaTreeview()
    Dim rsEditaTreeview As New ADODB.Recordset
    Dim SqlEditaTreeview As String
    vNmNo = Label8
    vNmNo = Replace(vNmNo, "/", ";")
    SqlEditaTreeview = "Select idform,nmform,formula,parametros from tbFormula where nmform = '" & vNmNo & "'"
    rsEditaTreeview.Open SqlEditaTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsEditaTreeview.EOF Then txtformula(4) = rsEditaTreeview.Fields(0)
End Sub

Private Sub TreeView3_Click()
    AlteraTreeview
    LimpaVariaveis
    LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
    compoeDadosLVs
    compoeControles
End Sub

Private Sub txtformula_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            If txtformula(0).Text = "" Then
                MsgBox "Selecione primeiro um CC - Centro de Custo"
                Exit Sub
            End If
            CarregaTxt "tbCCusto", "idprd", "S", "", "", txtformula(0), txtformula(1), 0, 1, txtformula(0), "S", txtformula(1)
            montaEstrutTreeview
            LimpaVariaveis
            LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
            compoeDadosLVs
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            preparaDados
            txtResultado = ""
            calculaValores 1
        End If
    End Select
End Sub

'vPosicao indica a posicao da formula
Private Sub localizaFormula(vNForm As Integer, vPosicao As Integer)
    Dim rsFormula As New ADODB.Recordset
    Dim SqlFormula As String
    SqlFormula = "select * from tbFormula as a where a.idprd = '" & txtformula(0) & "' and idform = '" & vNForm & "'"
    rsFormula.Open SqlFormula, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsFormula.EOF Then
        If vPosicao = 1 Then
            txtformula(7).Text = rsFormula.Fields(4) 'Formula 2
            txtformula(8).Text = rsFormula.Fields(3) 'Parametros 2
        ElseIf vPosicao = 2 Then
            txtformula(10).Text = rsFormula.Fields(4) 'Formula 2
            txtformula(9).Text = rsFormula.Fields(3) 'Parametros 2
        End If
    End If
    rsFormula.Close
    Set rsFormula = Nothing
End Sub

Private Sub substituiValores(vFormula As TextBox)
    Dim X As Integer
    Dim vPreserva As String
    vPreserva = ""
    vPreserva = vFormula
    For X = 1 To 50
        vFormula = Replace(vFormula, "cons(" & (X) & ")", cons(X))
        vFormula = Replace(vFormula, "var(" & (X) & ")", var(X))
        vFormula = Replace(vFormula, "vTMedio", vTMedio)
        vFormula = Replace(vFormula, "vFFadiga", vFFadiga)
        vFormula = Replace(vFormula, "vOrganiza", vOrganiza)
    Next
    vFormula = Replace(vFormula, ",", ".")
    txtDecoder = vFormula
    vFormula = vPreserva
End Sub

Private Sub calculaValores(vQual As Integer)
    'O ScriptControl é um componente. Ele interpreta e executa a formula/expressão numérica de um textbox
    If vQual = 1 Then
        txtResultado = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
    Else
        vGrupo = "1"
        vDimValor = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
        vDimValor = Replace(vDimValor, ",", ".")
        vDimValor = Replace(vDimValor, "(", "")
        vDimValor = Replace(vDimValor, ")", "")
        'MsgBox vResultFormula
    End If
End Sub

Private Sub preparaDados()
    LimpaVariaveis
    If txtformula(5) = "" Then
        MsgBox "Favor informar o campo: " & txtformula(5).Tag, vbInformation, "Atenção"
        txtformula(5).SetFocus
        Exit Sub
    End If
    'Calcula as formulas carregadas a partir das funções abaixo carregadas
    'a partir dos dados informados no campo de variáveis
    If Mid$(txtformula(2).Text, 1, 7) <> "formula" Then
        separaDadosVar txtformula(5)
        separaDadosPar txtformula(2)
        separaDadosCons
        localizaClassificacao
        substituiValores txtformula(3)
    Else
        If txtformula(7) <> "" Then
            'Acha o resultado referente a formula1
            separaDadosVar txtformula(5)
            separaDadosCons
            substituiValores txtformula(7)
            calculaValores 2
            localizaClassificacao
        End If
        
        If txtformula(10) <> "" Then
            'Acha o resultado referente a formula3
            separaDadosVar txtformula(5)
            separaDadosCons
            substituiValores txtformula(10)
            calculaValores 2
            localizaClassificacao
        End If
        
        vTMedio = Format(vSomaTempo, "#,##0.00;(#,##0.00)")
        'Pega o resultado das formulas 1 e 2 e aplica na formula3
        separaDadosVar txtformula(5)
        separaDadosPar txtformula(2)
        separaDadosCons
'        localizaClassificacao
        substituiValores txtformula(3)
    End If
End Sub

Private Sub transfDesenhosSel(llng_Contador As Integer, vTV As TreeView)
    Dim vNomeNo As String
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    
    
    If vTV.Nodes(llng_Contador).Checked = True Then
        vNomeNo = vTV.Nodes(llng_Contador).FullPath
    End If
    vNomeNo = Replace(vNomeNo, "\", ";")
    vJuntaNome = vNomeNo
    
    separaDadosTree vJuntaNome
    vNomeC = Right(vNomeC, 5)
    vCodLM = Mid$(vNomeC, 1, 2)
    vCodSeq = Mid$(vNomeC, 3, 3)
    
    cnBanco.BeginTrans
    
    If vAcumula = vNomeC And Label6 <> "-" Then
        cnBanco.CommitTrans
        Exit Sub
    Else
        vAcumula = vNomeC
    End If
    
    If vTV.Name = "TreeView1" Then
        If vCodLM <> "" And vCodSeq <> "" Then
            SqlTransf = "Insert into tbMPDesSel(fce,codlm,codseq) Values('" & Val(txtformula(12)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
            rsTransf.Open SqlTransf, cnBanco
        End If
    ElseIf vTV.Name = "TreeView2" Then
        SqlTransf = "Delete from tbMPDesSel where fce = '" & Val(txtformula(12)) & "' and codlm = '" & Val(vCodLM) & "' and codseq = '" & Val(vCodSeq) & "'"
        rsTransf.Open SqlTransf, cnBanco
    End If
    
    cnBanco.CommitTrans
End Sub

Private Sub SomaLV(LV As ListView)
    Dim X As Integer, Y As Integer, F As Integer
    Y = LV.ListItems.Count
    Dim somaTempo As Double
    somaTempo = 0
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then F = X
    Next
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        somaTempo = somaTempo + LV.SelectedItem.ListSubItems.Item(6)
    Next
    If somaTempo <> 0 Then
        Text2.Text = Format(somaTempo, "#,##00.00;(#,##0.00)")
        LV.ListItems.Item(F).Selected = True
    End If
End Sub


Private Sub txtformula_LostFocus(Index As Integer)
    Select Case Index
    Case 5
        preparaDados
        txtResultado = ""
        calculaValores 1
    End Select
End Sub
