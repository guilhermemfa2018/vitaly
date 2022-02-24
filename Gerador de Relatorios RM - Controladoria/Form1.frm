VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Relatórios RM - Rev.:2"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18765
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   18765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCG 
      Caption         =   "1.2.23 - NF de Compra para Estoque (ICMS)"
      Height          =   255
      Index           =   19
      Left            =   4800
      TabIndex        =   22
      Top             =   3240
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.Frame Frame7 
      Caption         =   "Configuração de conexão DB RM Sistemas "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   14280
      TabIndex        =   44
      Top             =   3720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   48
         Text            =   "vigamax"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Text            =   "sa"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   46
         Text            =   "CORPORERM_SOBRA"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Text            =   "SRV1002\CORPORERM"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
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
         Left            =   2400
         TabIndex        =   52
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
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
         TabIndex        =   51
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
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
         TabIndex        =   49
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
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
         Left            =   2400
         TabIndex        =   50
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame18 
      Height          =   3495
      Left            =   14280
      TabIndex        =   70
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame19 
         Caption         =   "Locais de Estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   4095
         Begin VB.CheckBox chkCG 
            Caption         =   "008 - Ferramental (Inativo)"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   3855
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "007 - Escritório (Inativo)"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   3855
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "004 - Estoque de Terceiros"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   3855
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "003 - Matéria-Prima"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   33
            Top             =   1080
            Width           =   3855
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "002 - Pátio - Produto Acabado (Inativo)"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   3735
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "001 - Almoxarifado Consimíveis Indiretos"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "DB última restauração "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14280
      TabIndex        =   65
      Top             =   6600
      Width           =   4335
      Begin VB.Label Label5 
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   69
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Hora:"
         Height          =   255
         Left            =   2400
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Período "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   64
      Top             =   360
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117702657
         CurrentDate     =   41502
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117702657
         CurrentDate     =   41502
      End
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Encargos e Salários "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame14 
      Height          =   7335
      Left            =   8880
      TabIndex        =   61
      Top             =   120
      Width           =   5295
      Begin VB.Frame Frame16 
         Caption         =   "Tipos de Documentos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   5055
         Begin MSComctlLib.ListView ListView1 
            Height          =   5775
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   10186
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Período "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   3495
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   285
            Left            =   1800
            TabIndex        =   28
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   117702657
            CurrentDate     =   41508
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   117702657
            CurrentDate     =   41508
         End
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   17760
      Top             =   6120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   18120
      Top             =   6120
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   1560
      Picture         =   "Form1.frx":3469A
      Style           =   1  'Graphical
      TabIndex        =   60
      Tag             =   "Restaurar dados"
      ToolTipText     =   "Restaurar dados"
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   840
      Picture         =   "Form1.frx":35364
      Style           =   1  'Graphical
      TabIndex        =   59
      Tag             =   "Backup"
      ToolTipText     =   "Backup"
      Top             =   7560
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   18120
      Top             =   5760
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   120
      Picture         =   "Form1.frx":367AE
      Style           =   1  'Graphical
      TabIndex        =   43
      Tag             =   "Gerar relatório"
      ToolTipText     =   "Gerar relatório"
      Top             =   7560
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Custo Gerencial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame6 
         Caption         =   "Produto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   42
         Top             =   5640
         Width           =   4095
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Código Custo Gerencial "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   4680
         Width           =   4095
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Movimentos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   4095
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.06 - CTRC - Conhecimento de Transporte"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.14 - Entrega de Material enviado por Terceiros"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   2520
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.13 - NF de Material de Terceiros p/ Ind."
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.12 - NF de Prestação de Serviço"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.08 - NF de Aplicação Direta"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.07 - NF de Compra para Estoque"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.04 - NF de Simples Remessa"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Value           =   1  'Checked
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Período "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3495
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   117702657
            CurrentDate     =   41499
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   117702657
            CurrentDate     =   41499
         End
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Centro de Custo - Saida de Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   4560
      TabIndex        =   38
      Top             =   120
      Width           =   4215
      Begin VB.Frame Frame13 
         Caption         =   "FCE Observação "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   57
         Top             =   6480
         Width           =   3975
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "FCE "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   56
         Top             =   5640
         Width           =   3975
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Produto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   4800
         Width           =   3975
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Código Centro de Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   54
         Top             =   3960
         Width           =   3975
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Movimentos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   3975
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.07 - NF de Compra para Estoque (02-05)"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.10 - Baixa de Material Reservado"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.06 - CTRC - Conhecimento de Transporte"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "2.2.22 - Requisição de Material"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.12 - NF de Prestação de Serviço"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "1.2.08 - NF de Aplicação Direta"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Value           =   1  'Checked
            Width           =   3735
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Aguarde enquanto são geradas as informações do relatório..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   7800
      Visible         =   0   'False
      Width           =   18495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim a As Single

Private Sub Command1_Click()
    'MsgBox "Rotina em desenvolvimento", vbInformation, "Backup"
    Form1.MousePointer = 11
    Label1.Visible = True
    Timer2.Enabled = True
End Sub

Private Sub Command3_Click()
    Conectar2
    'MsgBox "Rotina em desenvolvimento", vbInformation, "Backup"
    Form1.MousePointer = 11
    Label1.Visible = True
    Timer3.Enabled = True
End Sub

Private Sub Backup()
'On Error GoTo Err
    Dim rsBackRest As New ADODB.Recordset
    Dim sqlBackRest As String

    nomeDB = "CORPORERM"
    backupFile = "N'H:\usuarios\BackupRM\BkpDADOS.bak'"
    sqlBackRest = "BACKUP DATABASE [" & nomeDB & "] TO DISK = " & backupFile & " WITH NOFORMAT, INIT,  NAME = N'ZEUS-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
    cnBanco.CommandTimeout = 0
    rsBackRest.Open (sqlBackRest), cnBanco, adOpenStatic, adLockPessimistic
    MsgBox "Backup realizado com sucesso! O programa precisa ser reiniciado"
    Form1.MousePointer = 0
    Form1.Label1.Visible = False
    Exit Sub
Err:
    MsgBox "Erro na rotina de backup"
End Sub

Private Sub Restore()
'On Error GoTo Err
    Dim rsRest As New ADODB.Recordset
    Dim sqlRest As String

    nomeDB = "CORPORERM_SOBRA"
    backupFile = "N'H:\usuarios\BackupRM\BkpDADOS.bak'"
    sqlRest = "alter database " & nomeDB & " set single_user with rollback immediate;RESTORE DATABASE [" & nomeDB & "] FROM DISK = " & backupFile & " WITH  FILE = 1,  NOUNLOAD,  STATS = 10; alter database " & nomeDB & " set multi_user with rollback immediate"
    cnBanco2.CommandTimeout = 0
    cnBanco2.Execute sqlRest
    MsgBox "Restore realizado com sucesso! O programa precisa ser reiniciado"
    Form1.MousePointer = 0
    Form1.Label1.Visible = False
    Exit Sub
Err:
    MsgBox "Erro na rotina de backup"
End Sub

Private Sub Command2_Click()
    CriaDbTemp
    Form1.MousePointer = 11
    Label1.Visible = True
    Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    If vColigada = 1 Then
        Form1.Caption = Form1.Caption & " [VIGA] Gerador de Relatórios RM - Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    ElseIf vColigada = 5 Then
        Form1.Caption = Form1.Caption & " [VITALY] Gerador de Relatórios RM - Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    Else
        Form1.Caption = Form1.Caption & " [LUNA] Gerador de Relatórios RM - Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Deseja encerrar a aplicação", vbQuestion + vbYesNo, "Gerador") = vbYes Then
        Form1.Caption = ""
        cnBanco.Close
        Set cnBanco = Nothing
        End
    End If
End Sub

Private Sub Option4_Click()
    Bloq_Desbloq
End Sub

Private Sub Timer1_Timer()
    chamaRel
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Backup
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Restore
    Timer3.Enabled = False
End Sub

Private Sub CriaDbTemp()
    'On Error Resume Next
    Dim rsTbTemp As New ADODB.Recordset
    Dim sqlTbTemp As String
    
    Dim rsDropDb As New ADODB.Recordset
    Dim sqlDropDb As String
    
    'ABAIXO: CRIA TABELA TEMPORÁRIA COM O VALOR DE ENTRADA DE PRODUTOS
    sqlDropDb = "Drop table VALORENTTEMP"
    rsDropDb.Open sqlDropDb, cnBanco

    sqlTbTemp = "CREATE TABLE VALORENTTEMP (IDPRD int,PRECOUNIT numeric(15, 4),TPMOV varchar(10),NUMEROSEQUENCIAL smallint)"
    rsTbTemp.Open sqlTbTemp, cnBanco

    
    sqlTbTemp = "use corporerm_sobra; WITH tmp AS (SELECT x.*, ROW_NUMBER() OVER (PARTITION BY x.IDPRD ORDER BY x.dataemissao DESC) AS rn FROM TITMMOV as x inner join TMOV AS z " & _
                "ON x.IDMOV = z.IDMOV AND z.CODTMV IN('1.2.07')) " & _
                "INSERT INTO VALORENTTEMP (IDPRD, PRECOUNIT, TPMOV,NUMEROSEQUENCIAL) " & _
                "SELECT IDPRD, PRECOUNITARIO, A.CODTMV,NUMEROSEQUENCIAL FROM tmp inner join TMOV AS A ON A.IDMOV = TMP.IDMOV AND A.CODTMV IN('1.2.07') " & _
                "WHERE rn = 1 order by IDPRD;"

    rsTbTemp.Open (sqlTbTemp), cnBanco, adOpenStatic, adLockPessimistic

'End
'ABAIXO QUERY PARA RETORNAR MAIOR PRECO (DESATIVADA)
'    sqlTbTemp = "CREATE TABLE VALORENTTEMP (IDPRD int,PRECOUNIT numeric(15, 4),TPMOV varchar(10))" & _
'                "INSERT INTO VALORENTTEMP (IDPRD, PRECOUNIT, TPMOV) " & _
'               "SELECT C.IDPRD,MAX(B.PRECOUNITARIO) as PRECOUNITARIO,A.CODTMV FROM TMOV AS A INNER JOIN TITMMOV AS B ON A.CODCOLIGADA = B.CODCOLIGADA AND A.IDMOV = B.IDMOV AND " & _
'               "A.CODTMV IN('1.2.07') INNER JOIN TPRD AS C ON B.IDPRD = C.IDPRD AND A.CODCOLIGADA = C.CODCOLIGADA WHERE B.PRECOUNITARIO IS NOT NULL and B.PRECOUNITARIO > 0 GROUP BY C.IDPRD, A.CODTMV"
'    rsTbTemp.Open sqlTbTemp, cnBanco
    'rsTbTemp.Close
   'ACIMA: CRIA TABELA TEMPORÁRIA COM O VALOR DE ENTRADA DE PRODUTOS
End Sub

Private Sub chamaRel()
    If Option1.Value = True Then
        vDataFilter1 = DTPicker1
        vDataFilter2 = DTPicker2
        vCustos = Text1.Text
        vProduto = Text2.Text
        Compoe_Mov
        FCRCustoGerencial.Show 1
    ElseIf Option2.Value = True Then
        vDataFilter1 = DTPicker3
        vDataFilter2 = DTPicker4
        vCustos = Text3.Text
        vProduto = Text8.Text
        vFCECC = Text9.Text
        Compoe_Mov
        FCRCentroCusto.Show 1
    ElseIf Option3.Value = True Then
        vDataFilter1 = DTPicker5
        vDataFilter2 = DTPicker6
        Compoe_Mov
        FCREncargosSalarios.Show 1
    ElseIf Option4.Value = True Then
        Compoe_Mov
        FCREstoque.Show 1
    End If
End Sub

Private Sub Form_Load()
    DTPicker1 = CDate("01/01/" & Year(Date))
    DTPicker2 = CDate("31/12/" & Year(Date))
    DTPicker3 = CDate("01/01/" & Year(Date))
    DTPicker4 = CDate("31/12/" & Year(Date))
    DTPicker5 = CDate("01/01/" & Year(Date))
    DTPicker6 = CDate("31/12/" & Year(Date))
    listview_cabecalho
    Conectar
    Compoe_Listview
    buscaDadosRestore
End Sub

Private Sub buscaDadosRestore()
    Dim rsRestore As New ADODB.Recordset
    Dim sqlRestore As String
    sqlRestore = "DECLARE @dbname sysname, @days int SET @dbname = NULL SET @days = -7 SELECT TOP 1 rsh.destination_database_name AS [Database], rsh.user_name AS [Restored By],  " & _
                    "CASE WHEN rsh.restore_type = 'D' THEN 'Banco de Dados' WHEN rsh.restore_type = 'F' THEN 'Arquivo' WHEN rsh.restore_type = 'G' THEN 'Grupo de Arquivo' WHEN rsh.restore_type = 'I' THEN 'Diferencial' " & _
                    "WHEN rsh.restore_type = 'L' THEN 'Log' WHEN rsh.restore_type = 'V' THEN 'Verificação' WHEN rsh.restore_type = 'R' THEN 'Reversão' ELSE rsh.restore_type END AS [Restore Type], " & _
                    "CONVERT (VARCHAR, rsh.restore_date, 103) as [Restore Started],CONVERT(VARCHAR(5),rsh.restore_date,114) AS HORA,bmf.physical_device_name AS [Restored From],rf.destination_phys_name AS [Restored To] " & _
                    "FROM msdb.dbo.restorehistory rsh INNER JOIN msdb.dbo.backupset bs ON rsh.backup_set_id = bs.backup_set_id INNER JOIN msdb.dbo.restorefile rf ON rsh.restore_history_id = rf.restore_history_id " & _
                    "INNER JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id WHERE rsh.restore_date >= DATEADD(dd, ISNULL(@days, -30), GETDATE())AND destination_database_name = ISNULL(@dbname, destination_database_name) AND " & _
                    "rsh.destination_database_name = 'CORPORERM_Sobra' ORDER BY rsh.restore_history_id DESC"
    rsRestore.Open sqlRestore, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsRestore.EOF Then
        Label4.Caption = rsRestore.Fields(3)
        Label5.Caption = rsRestore.Fields(4)
    End If
    rsRestore.Close
    Set rsRestore = Nothing
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Cod.Doc.", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 1.4
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub Compoe_Listview()
    Dim rsEncargos As New ADODB.Recordset
    Dim sqlEncargos As String
    
    Dim ItemLst As ListItem
    Dim X As Integer
    
    ' Compoe Listview1
    sqlEncargos = "select A.CODLANC,A.DESCRICAO from PLANCFINANC AS A where a.CODCOLIGADA = '" & vColigada & "' ORDER BY A.CODLANC"
    sqlEncargos = "select A.CODLANC AS ID,A.LANFIN AS TIPO,A.CODTDO AS COD_DOC,A.DESCRICAO from PLANCFINANC AS A where a.CODCOLIGADA = '" & vColigada & "' ORDER BY A.CODTDO"
    
    rsEncargos.Open sqlEncargos, cnBanco, adOpenKeyset, adLockReadOnly
    X = 0
    While Not rsEncargos.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsEncargos.Fields(0), "000"))
        ItemLst.SubItems(1) = "" & Format(rsEncargos.Fields(1), "000")
        ItemLst.SubItems(2) = "" & Format(rsEncargos.Fields(2), "000")
        ItemLst.SubItems(3) = "" & rsEncargos.Fields(3)
        rsEncargos.MoveNext
        X = X + 1
    Wend
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
    rsEncargos.Close
    Set rsEncargos = Nothing
End Sub

Private Sub Compoe_Mov()
    Dim X As Integer
    vMovs = ""
    If Option1.Value = True Then
        For X = 0 To 6
            If vMovs = "" Then
                If chkCG(X).Value = 1 Then vMovs = vMovs & "'" & Mid(chkCG(X).Caption, 1, 6) & "'" Else vMovs = vMovs & "''"
            Else
                If chkCG(X).Value = 1 Then vMovs = vMovs & ",'" & Mid(chkCG(X).Caption, 1, 6) & "'" Else vMovs = vMovs & ",''"
            End If
        Next
    ElseIf Option2.Value = True Then
        For X = 7 To 11
            If vMovs = "" Then
                If chkCG(X).Value = 1 Then vMovs = vMovs & "'" & Mid(chkCG(X).Caption, 1, 6) & "'" Else vMovs = vMovs & "''"
            Else
                If chkCG(X).Value = 1 Then vMovs = vMovs & ",'" & Mid(chkCG(X).Caption, 1, 6) & "'" Else vMovs = vMovs & ",''"
            End If
        Next
        
        'BLOCO INSERIDO PARA CONSEDERAR O MOVIMENTO 1.2.07 NA SAIDA DE MATERIAIS (2 ~ 5)
        If vMovs = "" Then
            If chkCG(18).Value = 1 Then vMovs = vMovs & "'" & Mid(chkCG(18).Caption, 1, 6) & "'" Else: vMovs = vMovs & "''"
        Else
            If chkCG(18).Value = 1 Then vMovs = vMovs & ",'" & Mid(chkCG(18).Caption, 1, 6) & "'" Else: vMovs = vMovs & ",''"
        End If
        
        'BLOCO INSERIDO PARA CONSEDERAR O MOVIMENTO 1.2.23 NA SAIDA DE MATERIAIS (2 ~ 5)
        If vMovs = "" Then
            If chkCG(19).Value = 1 Then vMovs = vMovs & "'" & Mid(chkCG(19).Caption, 1, 6) & "'" Else: vMovs = vMovs & "''"
        Else
            If chkCG(19).Value = 1 Then vMovs = vMovs & ",'" & Mid(chkCG(19).Caption, 1, 6) & "'" Else: vMovs = vMovs & ",''"
        End If
        
    ElseIf Option3.Value = True Then
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            If vMovs = "" Then
                If ListView1.ListItems.Item(X).Checked = True Then
                    vMovs = vMovs & "" & ListView1.SelectedItem.ListSubItems.Item(2) & ""
                Else
                    vMovs = vMovs & "''"
                End If
            Else
                If ListView1.ListItems.Item(X).Checked = True Then
                    vMovs = vMovs & "," & ListView1.SelectedItem.ListSubItems.Item(2) & ""
                Else
                    vMovs = vMovs & ",''"
                End If
            End If
        Next
        'vMovs = "007,012,''"
    ElseIf Option4.Value = True Then
        For X = 12 To 17
            If vMovs = "" Then
                If chkCG(X).Value = 1 Then vMovs = vMovs & "'" & Mid(chkCG(X).Caption, 1, 3) & "'" Else vMovs = vMovs & "''"
            Else
                If chkCG(X).Value = 1 Then vMovs = vMovs & ",'" & Mid(chkCG(X).Caption, 1, 3) & "'" Else vMovs = vMovs & ",''"
            End If
        Next
    
    End If
End Sub


Private Sub Bloq_Desbloq()
    Dim X As Integer
    If Option1.Value = True Then
        'ATIVA CUSTO GERENCIAL
        Frame1.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Frame5.Enabled = True
        Frame6.Enabled = True
        For X = 0 To 6
            chkCG(X).Enabled = True
        Next
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
        
        'DESATIVA CENTRO DE CUSTO
        Frame8.Enabled = False
        Frame9.Enabled = False
        Frame10.Enabled = False
        Frame11.Enabled = False
        Frame12.Enabled = False
        Frame13.Enabled = False
        For X = 7 To 11
            chkCG(X).Enabled = False
        Next
        chkCG(18).Enabled = False
        chkCG(19).Enabled = False
        
        DTPicker3.Enabled = False
        DTPicker4.Enabled = False
        
        'DESATIVA ENCARGOS E SALARIOS
        Frame15.Enabled = False
        Frame16.Enabled = False
        DTPicker5.Enabled = False
        DTPicker6.Enabled = False
        ListView1.Enabled = False
        
        'DESATIVA ESTOQUE
        Frame18.Enabled = False
        Frame19.Enabled = False
        For X = 12 To 17
            chkCG(X).Enabled = False
        Next
    ElseIf Option2.Value = True Then
        'DESATIVA CUSTO GERENCIAL
        Frame1.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        Frame5.Enabled = False
        Frame6.Enabled = False
        For X = 0 To 6
            chkCG(X).Enabled = False
        Next
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
        
        'ATIVA CENTRO DE CUSTO
        Frame8.Enabled = True
        Frame9.Enabled = True
        Frame10.Enabled = True
        Frame11.Enabled = True
        Frame12.Enabled = True
        Frame13.Enabled = True
        For X = 7 To 11
            chkCG(X).Enabled = True
        Next
        chkCG(18).Enabled = True
        chkCG(19).Enabled = True
        
        DTPicker3.Enabled = True
        DTPicker4.Enabled = True
    
        'DESATIVA ENCARGOS E SALARIOS
        Frame15.Enabled = False
        Frame16.Enabled = False
        DTPicker5.Enabled = False
        DTPicker6.Enabled = False
        ListView1.Enabled = False
        
        'DESATIVA ESTOQUE
        Frame18.Enabled = False
        Frame19.Enabled = False
        For X = 12 To 17
            chkCG(X).Enabled = False
        Next
    
    ElseIf Option3.Value = True Then
        'DESATIVA CUSTO GERENCIAL
        Frame1.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        Frame5.Enabled = False
        Frame6.Enabled = False
        For X = 0 To 6
            chkCG(X).Enabled = False
        Next
        chkCG(18).Enabled = False
        chkCG(19).Enabled = False
        
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
        
        'ATIVA CENTRO DE CUSTO
        Frame8.Enabled = False
        Frame9.Enabled = False
        Frame10.Enabled = False
        Frame11.Enabled = False
        Frame12.Enabled = False
        Frame13.Enabled = False
        For X = 7 To 11
            chkCG(X).Enabled = False
        Next
        DTPicker3.Enabled = False
        DTPicker4.Enabled = False
    
        'DESATIVA ENCARGOS E SALARIOS
        Frame15.Enabled = True
        Frame16.Enabled = True
        DTPicker5.Enabled = True
        DTPicker6.Enabled = True
        ListView1.Enabled = True
        
        'DESATIVA ESTOQUE
        Frame18.Enabled = False
        Frame19.Enabled = False
        For X = 12 To 17
            chkCG(X).Enabled = False
        Next
        
    ElseIf Option4.Value = True Then
        'DESATIVA CUSTO GERENCIAL
        Frame1.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        Frame5.Enabled = False
        Frame6.Enabled = False
        For X = 0 To 6
            chkCG(X).Enabled = False
        Next
        chkCG(18).Enabled = False
        chkCG(19).Enabled = False
        
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
        
        'DESATIVA CENTRO DE CUSTO
        Frame8.Enabled = False
        Frame9.Enabled = False
        Frame10.Enabled = False
        Frame11.Enabled = False
        Frame12.Enabled = False
        Frame13.Enabled = False
        For X = 7 To 11
            chkCG(X).Enabled = False
        Next
        DTPicker3.Enabled = False
        DTPicker4.Enabled = False
        
        'DESATIVA ENCARGOS E SALARIOS
        Frame15.Enabled = False
        Frame16.Enabled = False
        DTPicker5.Enabled = False
        DTPicker6.Enabled = False
        ListView1.Enabled = False
    
        'ATIVA ESTOQUE
        Frame18.Enabled = True
        Frame19.Enabled = True
        For X = 12 To 17
            chkCG(X).Enabled = True
        Next
    
    
    End If
End Sub

Private Sub Option1_Click()
    Bloq_Desbloq
End Sub

Private Sub Option2_Click()
    Bloq_Desbloq
End Sub

Private Sub Option3_Click()
    Bloq_Desbloq
End Sub


