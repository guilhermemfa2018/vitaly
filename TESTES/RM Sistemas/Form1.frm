VERSION 5.00
Object = "{879115B9-8D7C-43CA-ADFE-8B489017BF42}#1.0#0"; "activelock1884.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor RM"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command3 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   77
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Frame Frame10 
      Caption         =   "Configurações "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   4080
      TabIndex        =   62
      Top             =   3360
      Width           =   6135
      Begin VB.Frame Frame13 
         Caption         =   "Senha de acesso "
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
         Left            =   120
         TabIndex        =   75
         Top             =   1560
         Width           =   2775
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
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   76
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Monitorar ociosidade:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   2775
         Begin VB.CheckBox Option1 
            Caption         =   "Sistema Operacional"
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
            TabIndex        =   74
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox Option2 
            Caption         =   "Módulos RM"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Teclas de atalho para modo de configuração "
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
         TabIndex        =   69
         Top             =   2280
         Width           =   5895
         Begin VB.ComboBox Combo1 
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
            ItemData        =   "Form1.frx":0CCA
            Left            =   1080
            List            =   "Form1.frx":0D7F
            TabIndex        =   71
            Text            =   "F12"
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label22 
            Caption         =   "Ctrl+Shift+"
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
            TabIndex        =   70
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Bloqueios"
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
         Left            =   3000
         TabIndex        =   64
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox Check9 
            Caption         =   "Regedit"
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
            TabIndex        =   68
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Prompt de comando"
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
            TabIndex        =   67
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Iniciar/Executar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Gerenciador de tarefas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Incializar no MSConfig"
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
         TabIndex        =   63
         Top             =   3240
         Width           =   2295
      End
   End
   Begin VB.ComboBox cboChild 
      Height          =   315
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Frame fraWindowInfo 
      Caption         =   "Informações do windows"
      Height          =   1455
      Left            =   10440
      TabIndex        =   54
      Top             =   2640
      Width           =   3735
      Begin VB.TextBox txtClassName 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txthWnd 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtWindowText 
         Height          =   285
         Left            =   720
         TabIndex        =   55
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblClassName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   960
         Width           =   420
      End
      Begin VB.Label lblhWnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hWnd:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Título:"
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   450
      End
   End
   Begin VB.ListBox lstWindow2 
      Height          =   3180
      Left            =   10440
      Sorted          =   -1  'True
      TabIndex        =   53
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1800
      Top             =   7560
   End
   Begin VB.ListBox lstWindow 
      Height          =   840
      ItemData        =   "Form1.frx":0E6F
      Left            =   5040
      List            =   "Form1.frx":0E71
      Sorted          =   -1  'True
      TabIndex        =   52
      Top             =   7080
      Width           =   5115
   End
   Begin VB.Frame Frame9 
      Caption         =   "Ociosidade Módulos RM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   3735
      Begin VB.TextBox Text12 
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
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Contábil"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text11 
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
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Pagamento"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text10 
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
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Fiscal"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text9 
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
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Financeiro"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text8 
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
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Estoque"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "RM Nucleus"
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
         Left            =   480
         TabIndex        =   41
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "RM Fluxus"
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
         Left            =   480
         TabIndex        =   40
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "RM Liber"
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
         Left            =   480
         TabIndex        =   39
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "RM Labore"
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
         Left            =   480
         TabIndex        =   38
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         Caption         =   "RM Saldus"
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
         Left            =   480
         TabIndex        =   37
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Image Image10 
         Height          =   270
         Left            =   120
         Picture         =   "Form1.frx":0E73
         Top             =   1800
         Width           =   270
      End
      Begin VB.Image Image9 
         Height          =   270
         Left            =   120
         Picture         =   "Form1.frx":11B6
         Top             =   1440
         Width           =   270
      End
      Begin VB.Image Image8 
         Height          =   270
         Left            =   120
         Picture         =   "Form1.frx":1504
         Top             =   1080
         Width           =   270
      End
      Begin VB.Image Image7 
         Height          =   270
         Left            =   120
         Picture         =   "Form1.frx":185A
         Top             =   360
         Width           =   270
      End
      Begin VB.Image Image6 
         Height          =   270
         Left            =   120
         Picture         =   "Form1.frx":1B96
         Top             =   720
         Width           =   270
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   120
         Picture         =   "Form1.frx":1ED5
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   120
         Picture         =   "Form1.frx":86B7
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "Form1.frx":EE99
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "Form1.frx":1567B
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "Form1.frx":1BE5D
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblStatusNucleus 
         Caption         =   "ocioso"
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
         Left            =   2880
         TabIndex        =   51
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label31 
         Caption         =   "-"
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
         Left            =   2160
         TabIndex        =   50
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label41 
         Caption         =   "-"
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
         Left            =   2160
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label51 
         Caption         =   "-"
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
         Left            =   2160
         TabIndex        =   48
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label61 
         Caption         =   "-"
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
         Left            =   2160
         TabIndex        =   47
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label71 
         Caption         =   "-"
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
         Left            =   2160
         TabIndex        =   46
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblStatusFluxus 
         Caption         =   "ocioso"
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
         Left            =   2880
         TabIndex        =   45
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblStatusLiber 
         Caption         =   "ocioso"
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
         Left            =   2880
         TabIndex        =   44
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblStatusLabore 
         Caption         =   "ocioso"
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
         Left            =   2880
         TabIndex        =   43
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblStatusSaldus 
         Caption         =   "ocioso"
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
         Left            =   2880
         TabIndex        =   42
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   7560
   End
   Begin VB.Frame Frame7 
      Caption         =   "Desconexões efetuadas "
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
      Left            =   120
      TabIndex        =   34
      Top             =   6840
      Width           =   3735
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Tempo limite de ociosidade/seg."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      TabIndex        =   32
      Top             =   2160
      Width           =   3015
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tempo limite de ociosidade/min."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      TabIndex        =   30
      Top             =   2160
      Width           =   3015
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   31
         Text            =   "01"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
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
      Left            =   3240
      TabIndex        =   29
      Tag             =   "Configurações"
      ToolTipText     =   "Configurações"
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
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
      Left            =   3960
      TabIndex        =   27
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame Frame4 
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
      Left            =   4080
      TabIndex        =   18
      Top             =   120
      Width           =   6135
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
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text3 
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
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text2 
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
         Left            =   3120
         TabIndex        =   22
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   2655
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
         Left            =   3120
         TabIndex        =   25
         Top             =   1200
         Width           =   2655
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
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   2655
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
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   2775
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
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estação "
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
      TabIndex        =   16
      Top             =   120
      Width           =   3135
      Begin VB.Label Label15 
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ociosidade do SO:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   3735
      Begin VB.Label Label14 
         Caption         =   "Última:"
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
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Atual:"
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
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuários por módulo: "
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
      TabIndex        =   0
      Top             =   4560
      Width           =   3735
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Contábil:"
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
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Pagamento:"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Fiscal:"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Financeiro:"
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
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Estoque:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   7560
   End
   Begin activelock1884.ActiveLock aLock 
      Left            =   1320
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   820
      SoftwareName    =   "monitortotvs"
      SoftwarePassword=   "2012"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
      LockToHardDrive =   0   'False
      LockToWindowsSerial=   -1  'True
      LockToRandomNumber=   -1  'True
      LockToComputerName=   0   'False
      LockToMACAddress=   0   'False
      UseDataLock     =   0   'False
      RegistryPath    =   "ActiveLock"
      RegistryKey     =   "VB and VBA Program Settings"
      RegistryName    =   "MyRegName"
      RegistryHive    =   "HKLM"
      LockToCustomString=   ""
      HashAlgorithm   =   0
      RegCounterKey   =   "Counter"
      RegLiberationKey=   "LiberationKey"
      RegLastRunDateKey=   "LastRunDate"
      RegInitialRunDateKey=   "InitialRunDate"
      RegRandomKey    =   "RandomKey"
      EncKey          =   "Default"
      RegEncKey       =   -1  'True
   End
   Begin VB.Image imgNoRegistrado 
      Height          =   555
      Left            =   3360
      Picture         =   "Form1.frx":2263F
      Tag             =   "Não registrado"
      ToolTipText     =   "Não registrado"
      Top             =   240
      Width           =   450
   End
   Begin VB.Image imgRegistrado 
      Height          =   570
      Left            =   3360
      Picture         =   "Form1.frx":22ABC
      Tag             =   "Registrado"
      ToolTipText     =   "Registrado"
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label20 
      Caption         =   "Falha na conexão com DB RM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   7680
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CONTROLA OCIOSIDADE DOS MODULOS RM
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Dim wnd As New Window
'REFERENTE A PARADA DE SERVIÇOS
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'REFERENTE A ÍCONE DO PROGRAMA NO SYSTRAY
Dim nid As NOTIFYICONDATA
Private Reg As Object
'VARIAVEIS DE CONTROLE DE TEMPO DOS MODULOS RM
Private contaResult As Integer
Private contaTempoNucleus As Integer
Private contaTempoFluxus As Integer
Private contaTempoLiber As Integer
Private contaTempoLabore As Integer
Private contaTempoSaldus As Integer
Private identKey As String
Private diasQueFaltaParaRegistrar As Integer


Private Sub Command1_Click()
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sServerName", Text1.Text 'Chave com o nome do Servidor
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sDatabaseName", Text2.Text 'Chave com o nome do Banco
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sUsuName", Text3.Text 'Chave com o usuario do banco
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sSenhaDB", Text4.Text 'Chave com Senha do banco
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sTimemin", Text5.Text 'Tempo limite de ociosidade em min
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sTimeseg", Text6.Text 'Tempo limite de ociosidade em seg
    If Option1.Value = 1 And Option2.Value = 0 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sMonitoraOQ", "SO"   'O que monitorar? Sist. Op. ou Mod. RM?
    If Option2.Value = 1 And Option1.Value = 0 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sMonitoraOQ", "RM"  'O que monitorar? Sist. Op. ou Mod. RM?
    If Option2.Value = 1 And Option1.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sMonitoraOQ", "TD"  'O que monitorar? Sist. Op. ou Mod. RM?
    
    If Check1.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMNucleus", "1" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMNucleus", "0"
    If Check2.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMFluxus", "1" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMFluxus", "0"
    If Check3.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMLiber", "1" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMLiber", "0"
    If Check4.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMLabore", "1" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMLabore", "0"
    If Check5.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMSaldus", "1" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRMSaldus", "0"
    
    If Check6.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sTaskManager", "S" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sTaskManager", "N" 'Gerenciador de tarefas
    If Check7.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sExecutar", "S" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sExecutar", "N" 'Executar
    If Check8.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sPrompt", "S" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sPrompt", "N" 'Prompt
    If Check9.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRegedit", "S" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sRegedit", "N" 'Regedit
    If Check10.Value = 1 Then Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sMSConfig", "S" Else Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sMSConfig", "N" 'Inicializar no MSConfig
    
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sPass", Text7.Text 'Senha de acesso
    encryptKey Combo1
    
    bloQueios
    
    carregaDados
    
    Form1.Width = 4035
    Command2.Caption = ">>"
    sDesativaCheckBox
End Sub

Private Sub bloQueios()
    'Bloqueia/desbloqueia Gerenciador de tarefas
    Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\" & "DisableTaskMgr", Check6.Value, "REG_DWORD"
    'Bloqueia/desbloqueia Inicia/Executar
    Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\" & "NoRun", Check7.Value, "REG_DWORD" 'Desativa o item "Executar" do menu Iniciar
    'Bloqueia/desbloqueia Prompt do MS-DOS
    Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\" & "Disabled", Check8.Value, "REG_DWORD" 'Desativa o Prompt do MS-DOS
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\" & "Disabled", Check8.Value, "REG_DWORD" 'Desativa o Prompt do MS-DOS
    'Bloqueia/desbloqueia Regedit
    Reg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\" & "DisableRegistryTools", Check9.Value, "REG_DWORD" 'Desativa regedit
    'inicializa no msconfig
    If Check10.Value = 1 Then
        Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & "MonitorTotvs", "C:\Arquivos de programas\MonitorTotvs\MonitorTotvs.exe"
    Else
        Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\MonitorTotvs"
    End If
End Sub

Private Function encryptKey(chave As String)
    If chave = "F1" Then chave = vbKeyF1
    If chave = "F2" Then chave = vbKeyF2
    If chave = "F3" Then chave = vbKeyF3
    If chave = "F4" Then chave = vbKeyF4
    If chave = "F5" Then chave = vbKeyF5
    If chave = "F6" Then chave = vbKeyF6
    If chave = "F7" Then chave = vbKeyF7
    If chave = "F8" Then chave = vbKeyF8
    If chave = "F9" Then chave = vbKeyF9
    If chave = "F10" Then chave = vbKeyF10
    If chave = "F11" Then chave = vbKeyF11
    If chave = "F12" Then chave = vbKeyF12
    If chave = "TAB" Then chave = vbKeyTab
    If chave = "ENTER" Then chave = vbKeyReturn
    If chave = "SPACEBAR" Then chave = vbKeySpace
    If chave = "PAGE UP" Then chave = vbKeyPageUp
    If chave = "PAGE DOWN" Then chave = vbKeyPageDown
    If chave = "END" Then chave = vbKeyEnd
    If chave = "HOME" Then chave = vbKeyHome
    If chave = "INSERT" Then chave = vbKeyInsert
    If chave = "DELETE" Then chave = vbKeyDelete
    If chave = "UP" Then chave = vbKeyUp
    If chave = "DOWN" Then chave = vbKeyDown
    If chave = "LEFT" Then chave = vbKeyLeft
    If chave = "RIGHT" Then chave = vbKeyRight
    If chave = "A" Then chave = vbKeyA
    If chave = "B" Then chave = vbKeyB
    If chave = "C" Then chave = vbKeyC
    If chave = "D" Then chave = vbKeyD
    If chave = "E" Then chave = vbKeyE
    If chave = "F" Then chave = vbKeyF
    If chave = "G" Then chave = vbKeyG
    If chave = "H" Then chave = vbKeyH
    If chave = "I" Then chave = vbKeyI
    If chave = "J" Then chave = vbKeyJ
    If chave = "K" Then chave = vbKeyK
    If chave = "L" Then chave = vbKeyL
    If chave = "M" Then chave = vbKeyM
    If chave = "N" Then chave = vbKeyN
    If chave = "O" Then chave = vbKeyO
    If chave = "P" Then chave = vbKeyP
    If chave = "Q" Then chave = vbKeyQ
    If chave = "R" Then chave = vbKeyR
    If chave = "S" Then chave = vbKeyS
    If chave = "T" Then chave = vbKeyT
    If chave = "U" Then chave = vbKeyU
    If chave = "V" Then chave = vbKeyV
    If chave = "W" Then chave = vbKeyW
    If chave = "X" Then chave = vbKeyX
    If chave = "Y" Then chave = vbKeyY
    If chave = "Z" Then chave = vbKeyZ
    If chave = "1" Then chave = vbKey1
    If chave = "2" Then chave = vbKey2
    If chave = "3" Then chave = vbKey3
    If chave = "4" Then chave = vbKey4
    If chave = "5" Then chave = vbKey5
    If chave = "6" Then chave = vbKey6
    If chave = "7" Then chave = vbKey7
    If chave = "8" Then chave = vbKey8
    If chave = "9" Then chave = vbKey9
    If chave = "0" Then chave = vbKey0
    If chave = "1 NK" Then chave = vbKeyNumpad1
    If chave = "2 NK" Then chave = vbKeyNumpad2
    If chave = "3 NK" Then chave = vbKeyNumpad3
    If chave = "4 NK" Then chave = vbKeyNumpad4
    If chave = "5 NK" Then chave = vbKeyNumpad5
    If chave = "6 NK" Then chave = vbKeyNumpad6
    If chave = "7 NK" Then chave = vbKeyNumpad7
    If chave = "8 NK" Then chave = vbKeyNumpad8
    If chave = "9 NK" Then chave = vbKeyNumpad9
    If chave = "0 NK" Then chave = vbKeyNumpad0
    If chave = "- NK" Then chave = vbKeySubtract
    If chave = "+ NK" Then chave = vbKeyAdd
    If chave = "* NK" Then chave = vbKeyMultiply
    If chave = "/ NK" Then chave = vbKeyDivide
    Reg.RegWrite "HKEY_CURRENT_USER\Software\MonitorRM\" & "sKey", chave 'Tecla de atalho
    identKey = chave
End Function

Private Sub decryptKey()
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF1 Then Combo1.Text = "F1"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF2 Then Combo1.Text = "F2"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF3 Then Combo1.Text = "F3"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF4 Then Combo1.Text = "F4"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF5 Then Combo1.Text = "F5"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF6 Then Combo1.Text = "F6"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF7 Then Combo1.Text = "F7"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF8 Then Combo1.Text = "F8"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF9 Then Combo1.Text = "F9"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF10 Then Combo1.Text = "F10"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF11 Then Combo1.Text = "F11"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF12 Then Combo1.Text = "F12"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyTab Then Combo1.Text = "TAB"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyReturn Then Combo1.Text = "ENTER"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeySpace Then Combo1.Text = "SPACEBAR"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyPageUp Then Combo1.Text = "PAGE UP"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyPageDown Then Combo1.Text = "PAGE DOWN"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyEnd Then Combo1.Text = "END"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyHome Then Combo1.Text = "HOME"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyInsert Then Combo1.Text = "INSERT"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyDelete Then Combo1.Text = "DELETE"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyUp Then Combo1.Text = "UP"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyDown Then Combo1.Text = "DOWN"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyLeft Then Combo1.Text = "LEFT"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyRight Then Combo1.Text = "RIGHT"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyA Then Combo1.Text = "A"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyB Then Combo1.Text = "B"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyC Then Combo1.Text = "C"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyD Then Combo1.Text = "D"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyE Then Combo1.Text = "E"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyF Then Combo1.Text = "F"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyG Then Combo1.Text = "G"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyH Then Combo1.Text = "H"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyI Then Combo1.Text = "I"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyJ Then Combo1.Text = "J"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyK Then Combo1.Text = "K"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyL Then Combo1.Text = "L"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyM Then Combo1.Text = "M"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyN Then Combo1.Text = "N"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyO Then Combo1.Text = "O"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyP Then Combo1.Text = "P"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyQ Then Combo1.Text = "Q"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyR Then Combo1.Text = "R"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyS Then Combo1.Text = "S"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyT Then Combo1.Text = "T"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyU Then Combo1.Text = "U"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyV Then Combo1.Text = "V"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyW Then Combo1.Text = "W"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyX Then Combo1.Text = "X"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyY Then Combo1.Text = "Y"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyZ Then Combo1.Text = "Z"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey1 Then Combo1.Text = "1"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey2 Then Combo1.Text = "2"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey3 Then Combo1.Text = "3"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey4 Then Combo1.Text = "4"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey5 Then Combo1.Text = "5"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey6 Then Combo1.Text = "6"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey7 Then Combo1.Text = "7"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey8 Then Combo1.Text = "8"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey9 Then Combo1.Text = "9"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKey0 Then Combo1.Text = "0"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad1 Then Combo1.Text = "1 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad2 Then Combo1.Text = "2 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad3 Then Combo1.Text = "3 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad4 Then Combo1.Text = "4 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad5 Then Combo1.Text = "5 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad6 Then Combo1.Text = "6 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad7 Then Combo1.Text = "7 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad8 Then Combo1.Text = "8 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad9 Then Combo1.Text = "9 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyNumpad0 Then Combo1.Text = "0 NK"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeySubtract Then Combo1.Text = "-"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyAdd Then Combo1.Text = "+"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyMultiply Then Combo1.Text = "*"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = vbKeyDivide Then Combo1.Text = "/"
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey") = "" Then Combo1.Text = "F12"
    identKey = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sKey")
End Sub

Private Sub Command2_Click()
    If Command2.Caption = ">>" Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Form1.Width = 10425
        Command2.Caption = "<<"
    Else
        carregaDados
        Form1.Width = 4035
        Command2.Caption = ">>"
    End If
    sDesativaCheckBox
End Sub

Private Sub listaModRM(sistem As String)
    Dim sTitle As String, hwnd As Long
    Dim TextoReal As String
    Dim TamanhoText1 As Integer, X As Integer
    hwnd = wnd.GetChildWindow(wnd.DesktopWindow)
    contaTempoNucleus = 0
    contaTempoFluxus = 0
    contaTempoLiber = 0
    contaTempoLabore = 0
    contaTempoSaldus = 0
    Do While hwnd <> 0
        wnd.hwnd = hwnd
        sTitle = wnd.GetText()
        TamanhoText1 = Len(sistem)
        TextoReal = Mid$(sistem, 1, TamanhoText1)
        If InStr(sTitle, TextoReal) > 0 Then
            If sTitle <> "" Then lstWindow.AddItem Trim(sTitle):  lstWindow.ItemData(lstWindow.NewIndex) = hwnd
            If sistem = "RM Nucleus" Then tempoOcioso sistem, contaTempoNucleus
            If sistem = "RM Fluxus" Then tempoOcioso sistem, contaTempoFluxus
            If sistem = "RM Liber" Then tempoOcioso sistem, contaTempoLiber
            If sistem = "RM Labore" Then tempoOcioso sistem, contaTempoLabore
            If sistem = "RM Saldus" Then tempoOcioso sistem, contaTempoSaldus
        End If
        hwnd = wnd.NextWindow(hwnd)
    Loop
    
    If contaTempoNucleus > 0 Then
        Label31 = Val(Label31) + 1
    ElseIf contaTempoNucleus = 0 And lblStatusNucleus = "ativo" Then
        Label31 = 0
    End If
    
    If contaTempoFluxus > 0 Then
        Label41 = Val(Label41) + 1 'Else Label41 = 0
    ElseIf contaTempoFluxus = 0 And lblStatusFluxus = "ativo" Then
        Label41 = 0
    End If
    
    If contaTempoLiber > 0 Then
        Label51 = Val(Label51) + 1 'Else Label51 = 0
    ElseIf contaTempoLiber = 0 And lblStatusLiber = "ativo" Then
        Label51 = 0
    End If
    If contaTempoLabore > 0 Then
        Label61 = Val(Label61) + 1 'Else Label61 = 0
    ElseIf contaTempoLabore = 0 And lblStatusLabore = "ativo" Then
        Label61 = 0
    End If
    If contaTempoSaldus > 0 Then
        Label71 = Val(Label71) + 1 'Else Label71 = 0
    ElseIf contaTempoSaldus = 0 And lblStatusSaldus = "ativo" Then
        Label71 = 0
    End If
   
    If Option1.Value = 0 And Option2.Value = 1 Then
        derrubaRM
    End If
    'QUANDO SELECIONA AS DUAS OPÇÕES, OBEDECE APENAS 60 SEGUNDOS DE OCIOSIDADE DO SISTEMA E NÃO O TEMPO
    'DETERMINADO PELO USUÁRIO
    If Option1.Value = 1 And Option2.Value = 1 Then
        If Val(Label1) >= 60 Or Val(Label31) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label41) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label51) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label61) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label71) >= Val(Text6) Then derrubaRM
    End If
End Sub

Private Sub Command3_Click()
    frmRegistro.Show 1
End Sub

Private Sub Form_Activate()
    If App.PrevInstance = True Then
        Form1.Visible = False
        Form1.Caption = "form2"
        winHWND = FindWindow(vbNullString, "Form1")
        SetActiveWindow winHWND
        AppActivate "Form1"
        SendKeys "% {Down 0} {Enter}", 1
        End
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Label3 = Label1
End Sub

Private Sub Form_Load()
    
    diasQueFaltaParaRegistrar = 0
    diasQueFaltaParaRegistrar = 30 - (aLock.UsedDays)
    
    If Not aLock.RegisteredUser Then
        imgNoRegistrado.Visible = True
        imgRegistrado.Visible = False
        
        Command3.Caption = "Você tem " & diasQueFaltaParaRegistrar & " dias - Registre-se"
        If diasQueFaltaParaRegistrar <= 0 Then
            frmRegistro.Show 1
            End
        End If
    Else
        Command3.Visible = False
        imgNoRegistrado.Visible = False
        imgRegistrado.Visible = True
    End If
    
    controleAcesso = 0
    carregaDados
    calcula
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Função q usa combinação de teclas para chamar outra função
    Dim TeclaSft        As Boolean
    Dim TeclaCtr        As Boolean
    TeclaSft = (Shift And vbShiftMask) > 0
    TeclaCtr = (Shift And vbCtrlMask) > 0
    If TeclaSft = False Or TeclaCtr = False Then Exit Sub
    If identKey = "" Then
        Command2_Click
    Else
        If TeclaSft And TeclaCtr And KeyCode = identKey Then
            If Form1.Width = 4035 Then
                Command2_Click
            End If
        End If
    End If
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If MsgBox("Tem a certeza que deseja fechar a aplicação?", vbQuestion + vbYesNo, "Projeto Desconecta RM") <> vbYes Then
        Cancel = True
     Else
        Cancel = True
        Form1.WindowState = 1
     End If
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Text5_LostFocus()
    calcula
End Sub

Private Sub calcula()
    Text6 = Val(Text5) * 60
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Label1 = fnIdleTime
    If Label3 > fnIdleTime = 0 Then
        Label3 = Label1
    End If
    
    'CONTROLA OCIOSIDADE SE APENAS OPTION1 ESTIVER MARCADO
    If Option1.Value = 1 And Option2.Value = 0 Then
        If Val(Label1) >= Val(Text6) Then
            fechaTOTVS
            Label21 = Val(Label21) + 1
        End If
    End If
    'CONTROLA OCIOSIDADE SE OPTION1 E OPTION2 ESTIVEREM MARCADOS
    If Option1.Value = 1 And Option2.Value = 1 Then
        'QUANDO SELECIONA AS DUAS OPÇÕES, OBEDECE APENAS 60 SEGUNDOS DE OCIOSIDADE DO SISTEMA E NÃO O TEMPO
        'DETERMINADO PELO USUÁRIO
        'If Val(Label1) >= 60 Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label31) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label41) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label51) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label61) >= Val(Text6) Then derrubaRM
        If Val(Label1) >= 60 Or Val(Label71) >= Val(Text6) Then derrubaRM
    End If
    Set Reg = Nothing
End Sub

Private Sub carregaDados()
    On Error Resume Next
    Set Reg = CreateObject("wscript.shell")
    
    senhaUsu = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sPass")
    Label9.Caption = Reg.RegRead("HKEY_CURRENT_USER\Software\RM Sistemas\RM Nucleus\LastUserName")
    Label10.Caption = Reg.RegRead("HKEY_CURRENT_USER\Software\RM Sistemas\RM Fluxus\LastUserName")
    Label11.Caption = Reg.RegRead("HKEY_CURRENT_USER\Software\RM Sistemas\RM Labore\LastUserName")
    Label12.Caption = Reg.RegRead("HKEY_CURRENT_USER\Software\RM Sistemas\RM Liber\LastUserName")
    Label13.Caption = Reg.RegRead("HKEY_CURRENT_USER\Software\RM Sistemas\RM Saldus\LastUserName")
    Label15.Caption = Reg.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\ComputerName\ComputerName\ComputerName")
    
    Text1.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sServerName") 'Chave com o nome do Servidor
    Text2.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sDatabaseName") 'Chave com o nome do Banco
    Text3.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sUsuName") 'Chave com o usuario do banco
    Text4.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sSenhaDB") 'Chave com Senha do banco
    
    Text5.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sTimemin") 'Chave com Senha do banco
    Text6.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sTimeseg") 'Chave com Senha do banco
    
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sMonitoraOQ") = "SO" Then Option1.Value = 1
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sMonitoraOQ") = "RM" Then Option2.Value = 1
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sMonitoraOQ") = "TD" Then
        Option1.Value = 1
        Option2.Value = 1
    End If
    
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sRMNucleus") = "1" Then Check1.Value = 1 Else Check1.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sRMFluxus") = "1" Then Check2.Value = 1 Else Check2.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sRMLiber") = "1" Then Check3.Value = 1 Else Check3.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sRMLabore") = "1" Then Check4.Value = 1 Else Check4.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sRMSaldus") = "1" Then Check5.Value = 1 Else Check5.Value = 0
    
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sTaskManager") = "S" Then Check6.Value = 1 Else Check6.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sExecutar") = "S" Then Check7.Value = 1 Else Check7.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sPrompt") = "S" Then Check8.Value = 1 Else Check8.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sRegedit") = "S" Then Check9.Value = 1 Else Check9.Value = 0
    If Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\sMSConfig") = "S" Then Check10.Value = 1 Else Check10.Value = 0
    Text7.Text = Reg.RegRead("HKEY_CURRENT_USER\Software\MonitorRM\\sPass") 'Senha de acesso ao programa
    
    decryptKey
    
    If Option1.Value = 1 And Option2.Value = 0 Then sAtivaDesativa "SO"
    If Option2.Value = 1 And Option1.Value = 0 Then sAtivaDesativa "RM"
    If Option1.Value = 1 And Option2.Value = 1 Then sAtivaDesativa "TD"
    
    If Option1.Value = 1 Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
    If Option2.Value = 1 Then
        Timer3.Enabled = True
    Else
        Timer3.Enabled = False
    End If
    
    If Label20.Visible = True Then
        Conectar
    End If
    sDesativaCheckBox
    Set Reg = Nothing
End Sub

'Rotinas abaixo Deixa o programa no SYSTRAY
Sub minimize_to_tray()
    Form1.Hide
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon
    nid.szTip = "Monitor de ociosidade TOTVS: " & App.Major & "." & App.Minor & "." & App.Revision & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
    Case WM_LBUTTONDBLCLK
    Me.Show
    Shell_NotifyIcon NIM_DELETE, nid
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = 0 And controleAcesso = 0 Then
        If senhaUsu <> "" Then
            controleAcesso = 1
            Form1.Visible = False
            Form2.Show 1
        End If
    End If
    
    If Me.WindowState = 1 Then
        minimize_to_tray
        Me.WindowState = 0
        controleAcesso = 0
    End If
End Sub

Private Sub fechaTOTVS()
On Error Resume Next
    updateGLOGIN
    Finalizar "RMNucleus"
    Finalizar "RMFluxus"
    Finalizar "RMLabore"
    Finalizar "rmsaldus"
    Finalizar "RMLiber"
    'PararServico
End Sub

Private Sub updateGLOGIN()
    Dim rsGLOGIN As New ADODB.Recordset
    Dim sqlGLOGIN As String
    
    Dim Y As Integer
    
    'Desconecta o usuário do RMNucleus
    sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label9.Caption & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'T'"
    rsGLOGIN.Open sqlGLOGIN, cnBanco
    
    'Desconecta o usuário do RMFluxus
    sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label10 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'F'"
    rsGLOGIN.Open sqlGLOGIN, cnBanco

    'Desconecta o usuário do RMLiber
    sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label11 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'D'"
    rsGLOGIN.Open sqlGLOGIN, cnBanco

    'Desconecta o usuário do RMLabore
    sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label12 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'P'"
    rsGLOGIN.Open sqlGLOGIN, cnBanco

    'Desconecta o usuário do RMSaldus
    sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label13 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'C'"
    rsGLOGIN.Open sqlGLOGIN, cnBanco
    
End Sub

Private Sub PararServico()
    Dim S As New UpcServiciosyWea
    On Error GoTo Mierda
    S.ServiceName = "RM.Host.Service"
    S.Stop_Service
    Sleep (2000)
    'Label22.Caption = "Parado"
Mierda:
End Sub

Private Sub IniciaServico()
    Dim S As New UpcServiciosyWea

    On Error GoTo Mierda
    S.ServiceName = "RM.Host.Service"

    S.Start_Service
    Sleep (2000)
    'Label22.Caption = "Inciado"
Mierda:
End Sub

Sub Finalizar(NomeExe As String)
    Dim IDProcesso As Long
    IDProcesso = GetProcessIDByEXEName(NomeExe)
    If Not IDProcesso = 0 Then ProcessTerminate IDProcesso
End Sub

Private Sub Timer3_Timer()
    lstWindow.Clear
    If Check1.Value = 1 Then listaModRM Check1.Caption Else: Label31 = "-"
    If Check2.Value = 1 Then listaModRM Check2.Caption Else: Label41 = "-"
    If Check3.Value = 1 Then listaModRM Check3.Caption Else: Label51 = "-"
    If Check4.Value = 1 Then listaModRM Check4.Caption Else: Label61 = "-"
    If Check5.Value = 1 Then listaModRM Check5.Caption Else: Label71 = "-"
    Static lHwnd As Long
    Dim lCurHwnd As Long
    Dim sText As String * 255
    lCurHwnd = GetForegroundWindow
    If lCurHwnd = lHwnd Then Exit Sub
    lHwnd = lCurHwnd
    If lHwnd <> hwnd Then
        Caption = "Janela Ativa: " & Left$(sText, GetWindowText(lHwnd, ByVal sText, 255))
        
'*ABAIXO TESTE *********************************
        If InStr(Caption, "RM Nucleus") = 0 Or InStr(Caption, "RM Fluxus") = 0 Or InStr(Caption, "RM Saldus") = 0 Or InStr(Caption, "RM Liber") = 0 Or InStr(Caption, "RM Labore") = 0 Then
            If InStr(Caption, "1.1.01") > 0 Then alterNameWindow "1.1.01"
            If InStr(Caption, "1.1.02") > 0 Then alterNameWindow "1.1.02"
            If InStr(Caption, "1.1.03") > 0 Then alterNameWindow "1.1.03"
            If InStr(Caption, "1.1.05") > 0 Then alterNameWindow "1.1.05"
            If InStr(Caption, "1.1.09") > 0 Then alterNameWindow "1.1.09"
            If InStr(Caption, "1.1.10") > 0 Then alterNameWindow "1.1.10"
            If InStr(Caption, "1.1.11") > 0 Then alterNameWindow "1.1.11"
            
            If InStr(Caption, "1.2.01") > 0 Then alterNameWindow "1.2.01"
            If InStr(Caption, "1.2.03") > 0 Then alterNameWindow "1.2.03"
            If InStr(Caption, "1.2.04") > 0 Then alterNameWindow "1.2.04"
            If InStr(Caption, "1.2.05") > 0 Then alterNameWindow "1.2.05"
            If InStr(Caption, "1.2.06") > 0 Then alterNameWindow "1.2.06"
            If InStr(Caption, "1.2.07") > 0 Then alterNameWindow "1.2.07"
            If InStr(Caption, "1.2.08") > 0 Then alterNameWindow "1.2.08"
            If InStr(Caption, "1.2.09") > 0 Then alterNameWindow "1.2.09"
            If InStr(Caption, "1.2.10") > 0 Then alterNameWindow "1.2.10"
            If InStr(Caption, "1.2.11") > 0 Then alterNameWindow "1.2.11"
            If InStr(Caption, "1.2.12") > 0 Then alterNameWindow "1.2.12"
            If InStr(Caption, "1.2.13") > 0 Then alterNameWindow "1.2.13"
            If InStr(Caption, "1.2.14") > 0 Then alterNameWindow "1.2.14"
            If InStr(Caption, "1.2.15") > 0 Then alterNameWindow "1.2.15"
            If InStr(Caption, "1.2.20") > 0 Then alterNameWindow "1.2.20"
        
            If InStr(Caption, "2.1.03") > 0 Then alterNameWindow "2.1.03"
            If InStr(Caption, "2.1.04") > 0 Then alterNameWindow "2.1.03"
            If InStr(Caption, "2.1.10") > 0 Then alterNameWindow "2.1.03"
            If InStr(Caption, "2.2.01") > 0 Then alterNameWindow "2.2.01"
            If InStr(Caption, "2.2.02") > 0 Then alterNameWindow "2.2.02"
            If InStr(Caption, "2.2.04") > 0 Then alterNameWindow "2.2.04"
            If InStr(Caption, "2.2.05") > 0 Then alterNameWindow "2.2.05"
            If InStr(Caption, "2.2.06") > 0 Then alterNameWindow "2.2.06"
            If InStr(Caption, "2.2.07") > 0 Then alterNameWindow "2.2.07"
            If InStr(Caption, "2.2.08") > 0 Then alterNameWindow "2.2.08"
            If InStr(Caption, "2.2.09") > 0 Then alterNameWindow "2.2.09"
            If InStr(Caption, "2.2.10") > 0 Then alterNameWindow "2.2.10"
            If InStr(Caption, "2.2.11") > 0 Then alterNameWindow "2.2.11"
            If InStr(Caption, "2.2.12") > 0 Then alterNameWindow "2.2.12"
            If InStr(Caption, "2.2.20") > 0 Then alterNameWindow "2.2.20"
            If InStr(Caption, "2.2.21") > 0 Then alterNameWindow "2.2.21"
            If InStr(Caption, "2.2.22") > 0 Then alterNameWindow "2.2.22"
        
            If InStr(Caption, "3.1.01") > 0 Then alterNameWindow "3.1.01"
            If InStr(Caption, "4.1.01") > 0 Then alterNameWindow "4.1.01"
        
            If InStr(Caption, "Produtos") > 0 Then alterNameWindow "Produtos"
            If InStr(Caption, "Local de Estoque") > 0 Then alterNameWindow "Local de Estoque"
            If InStr(Caption, "Cliente/Fornecedor:") > 0 Then alterNameWindow "Cliente/Fornecedor:"
            If InStr(Caption, "Usuário") > 0 Then alterNameWindow "Usuário"
            If InStr(Caption, "Vendedor/Comprador") > 0 Then alterNameWindow "Vendedor/Comprador"
        
            If InStr(Caption, "Lançamento:") > 0 Then alterNameWindow "Lançamento:"
            If InStr(Caption, "Extrato de Caixa:") > 0 Then alterNameWindow "Extrato de Caixa:"
            If InStr(Caption, "Conta/Caixa") > 0 Then alterNameWindow "Conta/Caixa"
            If InStr(Caption, "Tipos de Documento") > 0 Then alterNameWindow "Tipos de Documento"
        
            If InStr(Caption, "Plano de Contas") > 0 Then alterNameWindow "Plano de Contas"
            If InStr(Caption, "Lançamentos Contábeis") > 0 Then alterNameWindow "Lançamentos Contábeis"
        
            If InStr(Caption, "Lançamentos Entrada") > 0 Then alterNameWindow "Lançamentos Entrada"
            If InStr(Caption, "Lançamento Fiscal") > 0 Then alterNameWindow "Lançamento Fiscal"
            If InStr(Caption, "Lançamentos Saída") > 0 Then alterNameWindow "Lançamentos Saída"
            If InStr(Caption, "Período de Apuração") > 0 Then alterNameWindow "Período de Apuração"
            If InStr(Caption, "Lançamentos DIPAM") > 0 Then alterNameWindow "Lançamentos DIPAM"
            If InStr(Caption, "Ativo Imobilizado") > 0 Then alterNameWindow "Ativo Imobilizado"
            If InStr(Caption, "Natureza") > 0 Then alterNameWindow "Natureza"
        
            If InStr(Caption, "Evento:") > 0 Then alterNameWindow "Evento:"
                
        End If
'*ACIMA TESTE*********************************
        
    Else
        Caption = "Janela Ativa : Form1"
    End If
    percorreLst
End Sub

'************ TESTE************

Private Sub alterNameWindow(oQAlterar As String)
On Error Resume Next
    lstWindow2.Clear
    Dim sTitle As String, hwnd As Long
    Dim Y As Integer, X As Integer
    hwnd = wnd.GetChildWindow(wnd.DesktopWindow)
    Do While hwnd <> 0
        wnd.hwnd = hwnd
        sTitle = wnd.GetText()
        If sTitle <> "" Then
            If InStr(sTitle, oQAlterar) > 0 Then lstWindow2.AddItem Trim(sTitle):          lstWindow2.ItemData(lstWindow2.NewIndex) = hwnd
            'If sTitle = oQAlterar Then lstWindow2.AddItem Trim(sTitle):         lstWindow2.ItemData(lstWindow2.NewIndex) = hwnd
        End If
        hwnd = wnd.NextWindow(hwnd)
    Loop
    
    Y = lstWindow2.ListCount
    For X = 0 To Y
        lstWindow2.ListIndex = X
        If m_Picking = True Then Exit Sub
        wnd.hwnd = Val(txthWnd)
        
        If InStr(txtWindowText, "Janela Ativa") = 0 Then
        
        If InStr(Caption, "RM Nucleus") = 0 Then
            If oQAlterar = "1.1.01" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.1.02" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.1.03" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.1.05" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.1.09" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.1.10" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.1.11" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
        
            If oQAlterar = "1.2.01" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.02" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.03" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.04" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.05" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.06" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.07" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.08" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.09" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.10" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.11" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.12" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.13" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.14" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.15" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "1.2.20" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
        
            If oQAlterar = "2.1.03" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.1.04" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.1.10" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.01" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.02" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.04" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.05" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.06" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.07" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.08" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.09" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.10" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.11" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.12" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.20" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.21" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "2.2.22" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
    
            If oQAlterar = "3.1.01" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "4.1.01" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
    
            If oQAlterar = "Produtos" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "Local de Estoque" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "Cliente/Fornecedor:" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            
            If oQAlterar = "Usuário" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            If oQAlterar = "Vendedor/Comprador" Then wnd.SetText "RM Nucleus - " & txtWindowText.Text
            
        End If
        If InStr(Caption, "RM Fluxus") = 0 Then
            If oQAlterar = "Lançamento:" Then wnd.SetText "RM Fluxus - " & txtWindowText.Text
            If oQAlterar = "Extrato de Caixa:" Then wnd.SetText "RM Fluxus - " & txtWindowText.Text
            If oQAlterar = "Conta/Caixa" Then wnd.SetText "RM Fluxus - " & txtWindowText.Text
            If oQAlterar = "Tipos de Documento" Then wnd.SetText "RM Fluxus - " & txtWindowText.Text
        End If
        If InStr(Caption, "RM Saldus") = 0 Then
            If oQAlterar = "Plano de Contas" Then wnd.SetText "RM Saldus - " & txtWindowText.Text
            If oQAlterar = "Lançamentos Contábeis" Then wnd.SetText "RM Saldus - " & txtWindowText.Text
        End If
        If InStr(Caption, "RM Liber") = 0 Then
            If oQAlterar = "Lançamentos Entrada" Then wnd.SetText "RM Liber - " & txtWindowText.Text
            If oQAlterar = "Lançamento Fiscal" Then wnd.SetText "RM Liber - " & txtWindowText.Text
            If oQAlterar = "Lançamentos Saída" Then wnd.SetText "RM Liber - " & txtWindowText.Text
            If oQAlterar = "Período de Apuração" Then wnd.SetText "RM Liber - " & txtWindowText.Text
            If oQAlterar = "Lançamentos DIPAM" Then wnd.SetText "RM Liber - " & txtWindowText.Text
            If oQAlterar = "Ativo Imobilizado" Then wnd.SetText "RM Liber - " & txtWindowText.Text
            If oQAlterar = "Natureza" Then wnd.SetText "RM Liber - " & txtWindowText.Text
        End If
        If InStr(Caption, "RM Labore") = 0 Then
            If oQAlterar = "Evento:" Then wnd.SetText "RM Labore - " & txtWindowText.Text
        End If
    End If
        wnd.hwnd = Val(txthWnd)
        wnd.Flash
    Next
End Sub

'************ TESTE************
Private Sub lstWindow2_Click()
    cboChild.Clear
    Dim sTitle As String, hwnd As Long
    hwnd = wnd.GetChildWindow(lstWindow2.ItemData(lstWindow2.ListIndex))
    Do While hwnd <> 0
        wnd.hwnd = hwnd
        sTitle = wnd.GetText()
        If sTitle <> "" Then cboChild.AddItem Trim(sTitle): cboChild.ItemData(cboChild.NewIndex) = hwnd
        hwnd = wnd.NextWindow(hwnd)
    Loop
    If cboChild.ListCount > 0 Then cboChild.ListIndex = 0
    FillProps lstWindow2.ItemData(lstWindow2.ListIndex)
End Sub

'************ TESTE************
Sub FillProps(lWindow As Long)
    Dim sClassName As String
    m_hWnd = lWindow
    wnd.hwnd = lWindow
    txtWindowText.Text = wnd.GetText()
    txthWnd.Text = m_hWnd
    'Get Window's Class Name
    txtClassName.Text = wnd.GetWindowClassName()
    'Get Window Position
End Sub

Private Sub derrubaRM()
    Dim rsGLOGIN As New ADODB.Recordset
    Dim sqlGLOGIN As String
    Dim Y As Integer
    
    If Val(Label31) >= Val(Text6) And lstWindow.ListCount > 0 Or Val(Label1) >= Val(Text6) And lstWindow.ListCount > 0 Then
        Finalizar "RMNucleus"
        Label21 = Val(Label21) + 1
        'Desconecta o usuário do RMNucleus
        sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label9.Caption & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'T'"
        rsGLOGIN.Open sqlGLOGIN, cnBanco
        Label31 = 0
        Label1 = 0
    End If
    If Val(Label41) >= Val(Text6) And lstWindow.ListCount > 0 Or Val(Label1) >= Val(Text6) And lstWindow.ListCount > 0 Then
        Finalizar "RMFluxus"
        Label21 = Val(Label21) + 1
        'Desconecta o usuário do RMFluxus
        sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label10 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'F'"
        rsGLOGIN.Open sqlGLOGIN, cnBanco
        Label41 = 0
        Label1 = 0
    End If
    
    If Val(Label51) >= Val(Text6) And lstWindow.ListCount > 0 Or Val(Label1) >= Val(Text6) And lstWindow.ListCount > 0 Then
        Finalizar "RMLiber"
        Label21 = Val(Label21) + 1
        'Desconecta o usuário do RMLiber
        sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label11 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'D'"
        rsGLOGIN.Open sqlGLOGIN, cnBanco
        Label51 = 0
        Label1 = 0
    End If
    If Val(Label61) >= Val(Text6) And lstWindow.ListCount > 0 Or Val(Label1) >= Val(Text6) And lstWindow.ListCount > 0 Then
        Finalizar "RMLabore"
        Label21 = Val(Label21) + 1
        'Desconecta o usuário do RMLabore
        sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label12 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'P'"
        rsGLOGIN.Open sqlGLOGIN, cnBanco
        Label61 = 0
        Label1 = 0
    End If
    If Val(Label71) > Val(Text6) And lstWindow.ListCount > 0 Or Val(Label1) >= Val(Text6) And lstWindow.ListCount > 0 Then
        Finalizar "RMsaldus"
        Label21 = Val(Label21) + 1
        'Desconecta o usuário do RMSaldus
        sqlGLOGIN = "Delete from GLOGIN where GLOGIN.USERNAME = '" & Label13 & "' and GLOGIN.COMPUTERNAME = '" & Label15.Caption & "' and GLOGIN.CODSISTEMA = 'C'"
        rsGLOGIN.Open sqlGLOGIN, cnBanco
        Label71 = 0
        Label1 = 0
    End If
End Sub

Private Sub pesquisaString(ondeProcurar As String)
    Dim oQProcurar(5) As String
    Dim TamanhoText1 As Integer, X As Integer
    oQProcurar(1) = "RM Nucleus"
    oQProcurar(2) = "RM Fluxus"
    oQProcurar(3) = "RM Liber"
    oQProcurar(4) = "RM Labore"
    oQProcurar(5) = "RM Saldus"
    For X = 1 To 5
        TamanhoText1 = Len(oQProcurar(X))
        TextoReal = Mid$(oQProcurar(X), 1, TamanhoText1)
        If InStr(Caption, TextoReal) > 0 Then
            If TextoReal = "RM Nucleus" Then contaTempoNucleus = contaTempoNucleus + 1
            If TextoReal = "RM Fluxus" Then contaTempoFluxus = contaTempoFluxus + 1
            If TextoReal = "RM Liber" Then contaTempoLiber = contaTempoLiber + 1
            If TextoReal = "RM Labore" Then contaTempoLabore = contaTempoLabore + 1
            If TextoReal = "RM Saldus" Then contaTempoSaldus = contaTempoSaldus + 1
        End If
    Next
End Sub

Private Sub percorreLst()
    Y = lstWindow.ListCount
    contaTempoNucleus = 0
    contaTempoFluxus = 0
    contaTempoLiber = 0
    contaTempoLabore = 0
    contaTempoSaldus = 0
    For X = 1 To Y - 1
        Label2 = lstWindow.List(X)
        pesquisaString lstWindow.List(X)
    Next
    If contaTempoNucleus > 0 Then lblStatusNucleus.Caption = "ativo" Else lblStatusNucleus.Caption = "ocioso"
    If contaTempoFluxus > 0 Then lblStatusFluxus.Caption = "ativo" Else lblStatusFluxus.Caption = "ocioso"
    If contaTempoLiber > 0 Then lblStatusLiber.Caption = "ativo" Else lblStatusLiber.Caption = "ocioso"
    If contaTempoLabore > 0 Then lblStatusLabore.Caption = "ativo" Else lblStatusLabore.Caption = "ocioso"
    If contaTempoSaldus > 0 Then lblStatusSaldus.Caption = "ativo" Else lblStatusSaldus.Caption = "ocioso"
End Sub

Private Sub tempoOcioso(sistem As String, contaPassagem As Integer)
    If sistem = "RM Nucleus" And lblStatusNucleus.Caption = "ocioso" Then
        contaTempoNucleus = contaTempoNucleus + 1
    End If
    If sistem = "RM Fluxus" And lblStatusFluxus.Caption = "ocioso" Then
        contaTempoFluxus = contaTempoFluxus + 1
    End If
    If sistem = "RM Liber" And lblStatusLiber.Caption = "ocioso" Then
        contaTempoLiber = contaTempoLiber + 1
    End If
    If sistem = "RM Labore" And lblStatusLabore.Caption = "ocioso" Then
        contaTempoLabore = contaTempoLabore + 1
    End If
    If sistem = "RM Saldus" And lblStatusSaldus.Caption = "ocioso" Then
        contaTempoSaldus = contaTempoSaldus + 1
    End If
End Sub

Private Sub sDesativaCheckBox()
    If Command2.Caption = ">>" Then
        Check1.Enabled = False
        Check2.Enabled = False
        Check3.Enabled = False
        Check4.Enabled = False
        Check5.Enabled = False
    Else
        Check1.Enabled = True
        Check2.Enabled = True
        Check3.Enabled = True
        Check4.Enabled = True
        Check5.Enabled = True
    End If
End Sub
Private Sub sAtivaDesativa(oQ As String)
    If oQ = "SO" Then
        Frame2.Enabled = True
        Label2.Enabled = True
        Label14.Enabled = True
        Label1.Enabled = True
        Label3.Enabled = True
        
        Frame9.Enabled = False
        Check1.Enabled = False
        Check2.Enabled = False
        Check3.Enabled = False
        Check4.Enabled = False
        Check5.Enabled = False
        Label31.Enabled = False
        Label41.Enabled = False
        Label51.Enabled = False
        Label61.Enabled = False
        Label71.Enabled = False
        lblStatusNucleus.Enabled = False
        lblStatusFluxus.Enabled = False
        lblStatusLiber.Enabled = False
        lblStatusLabore.Enabled = False
        lblStatusSaldus.Enabled = False
    ElseIf oQ = "RM" Then
        Frame2.Enabled = False
        Label2.Enabled = False
        Label14.Enabled = False
        Label1.Enabled = False
        Label3.Enabled = False
        
        Frame9.Enabled = True
        Check1.Enabled = True
        Check2.Enabled = True
        Check3.Enabled = True
        Check4.Enabled = True
        Check5.Enabled = True
        Label31.Enabled = True
        Label41.Enabled = True
        Label51.Enabled = True
        Label61.Enabled = True
        Label71.Enabled = True
        lblStatusNucleus.Enabled = True
        lblStatusFluxus.Enabled = True
        lblStatusLiber.Enabled = True
        lblStatusLabore.Enabled = True
        lblStatusSaldus.Enabled = True
    ElseIf oQ = "RM" Then
        Frame2.Enabled = True
        Label2.Enabled = True
        Label14.Enabled = True
        Label1.Enabled = True
        Label3.Enabled = True
        
        Frame9.Enabled = True
        Check1.Enabled = True
        Check2.Enabled = True
        Check3.Enabled = True
        Check4.Enabled = True
        Check5.Enabled = True
        Label31.Enabled = True
        Label41.Enabled = True
        Label51.Enabled = True
        Label61.Enabled = True
        Label71.Enabled = True
        lblStatusNucleus.Enabled = True
        lblStatusFluxus.Enabled = True
        lblStatusLiber.Enabled = True
        lblStatusLabore.Enabled = True
        lblStatusSaldus.Enabled = True
    End If
        
End Sub
