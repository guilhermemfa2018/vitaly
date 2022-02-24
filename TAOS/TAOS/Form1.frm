VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Apropriação"
   ClientHeight    =   11280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1200
   End
   Begin VB.Frame Frame9 
      Caption         =   "Mapa de Registros das OS's em andamento"
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
      Left            =   360
      TabIndex        =   43
      Top             =   5640
      Width           =   14535
      Begin MSComctlLib.ListView ListView7 
         Height          =   3495
         Left            =   10920
         TabIndex        =   47
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   3495
         Left            =   7320
         TabIndex        =   46
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3495
         Left            =   3720
         TabIndex        =   45
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3495
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6165
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483638
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   14280
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame100 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   14895
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
         Left            =   6480
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Text            =   "SRV1002\CORPORERM"
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3120
            TabIndex        =   4
            Text            =   "ZEUS"
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Text            =   "sa"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox Text4 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3120
            PasswordChar    =   "*"
            TabIndex        =   2
            Text            =   "vigamax"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label17 
            Caption         =   "Nome do BANCO:"
            Height          =   255
            Left            =   3120
            TabIndex        =   9
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label16 
            Caption         =   "Nome do SERVIDOR:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label18 
            Caption         =   "USUÁRIO:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label19 
            Caption         =   "SENHA:"
            Height          =   255
            Left            =   3120
            TabIndex        =   6
            Top             =   1200
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Código de parada"
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
         Left            =   240
         TabIndex        =   22
         Top             =   4080
         Visible         =   0   'False
         Width           =   14535
         Begin VB.Frame Frame8 
            Caption         =   "Tipo OS"
            Height          =   735
            Left            =   11640
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
            Begin VB.Label lblTipoOS 
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   360
               Visible         =   0   'False
               Width           =   2415
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Paradas "
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
            Left            =   240
            TabIndex        =   23
            Top             =   1320
            Width           =   13215
            Begin MSComctlLib.ListView ListView3 
               Height          =   2055
               Left            =   8760
               TabIndex        =   24
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   3625
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483638
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   2055
               Left            =   4440
               TabIndex        =   25
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   3625
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483638
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               NumItems        =   0
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   2055
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   3625
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483638
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               NumItems        =   0
            End
         End
         Begin VB.TextBox txtOS 
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
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Não esqueça de registrar SAIDA ao final da parada"
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
            Height          =   255
            Left            =   2880
            TabIndex        =   57
            Top             =   120
            Visible         =   0   'False
            Width           =   11415
         End
         Begin VB.Label Label13 
            Caption         =   "Parada nº:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   1920
            Picture         =   "Form1.frx":1994
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label14 
            Caption         =   "Digite o código da PARADA via teclado no campo ao lado ou passe o código de barra da PARADA no leitor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   29
            Top             =   600
            Width           =   11775
         End
         Begin VB.Image Image4 
            Height          =   105
            Left            =   240
            Picture         =   "Form1.frx":265E
            Top             =   1200
            Width           =   11985
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "* Digite 0 para Cancelar o Procedimento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   2520
            TabIndex        =   28
            Top             =   840
            Width           =   9735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "OS - Ordem de Serviço "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   34
         Top             =   4080
         Visible         =   0   'False
         Width           =   14415
         Begin VB.TextBox txtOS 
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
            Height          =   480
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Caption         =   "* Digite 1000 para serviços sem OS "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   9480
            TabIndex        =   56
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Digite o código da OS via teclado no campo ao lado ou passe o código de barras da PARADA no leitor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   37
            Top             =   600
            Width           =   9255
         End
         Begin VB.Label Label11 
            Caption         =   "OS nº:"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   4200
            Picture         =   "Form1.frx":2E8E
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "* Digite 0 para Cancelar o Procedimento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   5280
            TabIndex        =   36
            Top             =   840
            Width           =   3975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Informações"
         Height          =   1575
         Left            =   240
         TabIndex        =   31
         Top             =   7920
         Width           =   14535
         Begin VB.Frame Frame6 
            Height          =   855
            Left            =   13560
            TabIndex        =   32
            Top             =   240
            Width           =   735
            Begin VB.Image Nok 
               Height          =   480
               Left            =   120
               Picture         =   "Form1.frx":3B58
               Top             =   240
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Image Ok 
               Height          =   480
               Left            =   120
               Picture         =   "Form1.frx":4822
               Top             =   240
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   13215
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   3495
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   14535
         Begin VB.Frame Frame11 
            Caption         =   "Hora do Sistema"
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
            Left            =   12480
            TabIndex        =   58
            Top             =   240
            Width           =   1935
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               Caption         =   "00:00:00"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Cronômetro "
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
            Left            =   10920
            TabIndex        =   48
            Top             =   2400
            Visible         =   0   'False
            Width           =   3495
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   1320
               Top             =   480
            End
            Begin VB.Label Horas 
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label20 
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   480
               TabIndex        =   52
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Label15 
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               TabIndex        =   51
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Minutos 
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   600
               TabIndex        =   50
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Segundos 
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1080
               TabIndex        =   49
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.TextBox txtOS 
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
            Height          =   450
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Registro nº:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Nome:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblOS 
            Caption         =   "NOME"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   12495
         End
         Begin VB.Label Label4 
            Caption         =   "Setor:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Função:"
            Height          =   255
            Left            =   5640
            TabIndex        =   17
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Centro de Custo:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lblOS 
            Caption         =   "SETOR"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   5175
         End
         Begin VB.Label lblOS 
            Caption         =   "FUNÇÃO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   5640
            TabIndex        =   14
            Top             =   2160
            Width           =   6975
         End
         Begin VB.Label lblOS 
            Caption         =   "CENTRO DE CUSTO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   2880
            Width           =   12495
         End
         Begin VB.Label Label10 
            Caption         =   "Digite seu registro via teclado no campo ao lado ou passe seu cracha no leitor de código de barras"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   12
            Top             =   720
            Width           =   9975
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3240
            Picture         =   "Form1.frx":54EC
            Top             =   600
            Width           =   480
         End
      End
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Label23"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   55
      Top             =   120
      Width           =   12135
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   54
      Top             =   1200
      Width           =   12015
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   960
      Left            =   120
      Top             =   240
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1693
      Image           =   "Form1.frx":61B6
      Props           =   5
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Terminal de Apropriação de Ordem de Serviço"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   39
      Top             =   480
      Width           =   12015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vSetFocus As Integer
Private vCBarraGeral As String
Private vFCEGlobal As Integer
Private vDesenhosGlobal As String

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Activate()
    txtOS(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    AlwaysOnTop Me, True ' Mantem o formulário sempre em primeiro plano
    Conectar
    Label23.Caption = Text6.Text
    Label22.Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    msgLabel "", 1
    
    carregaDadosEmail
    
    listview_cabecalho
    LimpaLV ListView1
    chamaSQL "select a.codigo,a.nmparada from tbParadas as a where a.codigo >=9001 and a.codigo <= 9007"
    Compoe_Listview ListView1, Sqlp, "0000"
    
    LimpaLV ListView2
    chamaSQL "select a.codigo,a.nmparada from tbParadas as a where a.codigo >=9008 and a.codigo <= 9014"
    Compoe_Listview ListView2, Sqlp, "0000"
    
    LimpaLV ListView3
    chamaSQL "select a.codigo,a.nmparada from tbParadas as a where a.codigo >=9015 and a.codigo <= 9023"
    Compoe_Listview ListView3, Sqlp, "0000"
    MapaOs
    lblOS(0) = "-"
    lblOS(1) = "-"
    lblOS(2) = "-"
    lblOS(3) = "-"
    vVerificaPermissao = 0
    Timer1.Enabled = True
End Sub

Private Sub txtOS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Err
    Select Case Index
    Case 0
        Frame3.Caption = "Código de parada "
        Label25.Visible = False
        
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If compoeControles = False Then Exit Sub
            
            If verificaSeApropria = False Then
                msgLabel "COLABORADOR NÃO TEM PERMISSÃO PARA APROPRIAR HORAS", 2
                txtOS(0).Text = ""
                Exit Sub
            End If
            'Verifica se se o COLABORADOR esta na tabela de PERMISSÕES
                        
            vVerificaPermissao = 0
            If VerificaPermissao = True Then
                'MsgBox "Colaborador com permissão de somente FECHAR OS", vbCritical, "TESTE"
                
                vVerificaPermissao = 1
                Frame2.Visible = True
                txtOS(1).SetFocus
                
                'End
            End If
            
            procuraSaida 'Verifica se o colaborador tem premissão de saida do sistema
            achaEntrada 'Habilita Listview de: ENTRADA ou de PARADA
            If Frame2.Visible = True Then
                txtOS(1).SetFocus
            ElseIf Frame3.Visible = True Then
                txtOS(2).SetFocus
            End If
            Frame9.Visible = False
        End If
    Case 1
        'REGISTRA OS
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtOS(1).Text = "0" Then
                msgLabel "PROCEDIMENTO CANCELADO", 2
                lblOS(0).Caption = "-"
                lblOS(1).Caption = "-"
                lblOS(2).Caption = "-"
                lblOS(3).Caption = "-"
                txtOS(0).Text = ""
                txtOS(1).Text = ""
                txtOS(2).Text = ""
                Frame2.Visible = False
                txtOS(0).SetFocus
                MapaOs
                Frame9.Visible = True
            ElseIf txtOS(1).Text = "1000" Then
                'UTILIZA A ROTINA DE PARADA PARA REGISTRAR SERVIÇOS EXTRA OS NO CODIGO 1000
                Frame3.Caption = "registrar ENTRADA em parada"
                Label25.Visible = True
                Frame3.Visible = True
                txtOS(2).SetFocus
            Else
                achaOS
            End If
            'MapaOs
            'Frame9.Visible = True
        End If
    Case 2
        'INICIA PROCEDIMENTO DE PARADA
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtOS(2).Text = "1" Then
                msgLabel "Código de PARADA não encontrado", 2
                txtOS(2).Text = ""
                Exit Sub
            End If
            
            If txtOS(2).Text = "0" Then
                msgLabel "PROCEDIMENTO CANCELADO", 2
            Else
                If achaParada = False Then Exit Sub
                If lblTipoOS = "Fechamento" Then
                    If validaCracha = False Then
                        msgLabel "Colaborador: " & lblOS(0) & " não tem permissão para FECHAR a operação: " & vCBarraGeral, 2
                        txtOS(2).Text = ""
                        txtOS(0).Text = ""
                        lblOS(0).Caption = "-"
                        lblOS(1).Caption = "-"
                        lblOS(2).Caption = "-"
                        lblOS(3).Caption = "-"
                        txtOS(0).Text = ""
                        txtOS(1).Text = ""
                        txtOS(2).Text = ""
                        Frame3.Visible = False
                        txtOS(0).SetFocus
                        MapaOs
                        Frame9.Visible = True
                        Exit Sub
                    End If
                    
                    'ROTINA DE FECHAMENTO DE OS
                    frmConfirmaChapa.Show 1
                    If vNomeGlobal = lblOS(0).Caption And vMSGGlobal = "OK" Then
                        fechaOS
                    Else
                        msgLabel "Colaborador: " & vNomeGlobal & " não tem permissão para FECHAR a operação: " & vCBarraGeral, 2
                    End If
                    'gravaParada2
                Else
                    If txtOS(1).Text <> "1000" Then
                        gravaParada
                        msgLabel "Parada nº: " & txtOS(2).Text & " - Registrada para: " & lblOS(0).Caption, 1
                    Else
                        'OS esta em aberto
                        msgLabel lblOS(0).Caption & " registrou ENTRADA para o serviço nº: " & txtOS(2).Text, 1
                        gravaOS
                        iniciaRetrabalho txtOS(1).Text
                        lblOS(0).Caption = "-"
                        lblOS(1).Caption = "-"
                        lblOS(2).Caption = "-"
                        lblOS(3).Caption = "-"
                        txtOS(0).Text = ""
                        txtOS(1).Text = ""
                        txtOS(2).Text = ""
                        Frame2.Visible = False
                        txtOS(0).SetFocus
                        MapaOs
                        Frame9.Visible = True
                    End If
                End If
                MapaOs
                Frame9.Visible = True
            End If
            lblOS(0).Caption = "-"
            lblOS(1).Caption = "-"
            lblOS(2).Caption = "-"
            lblOS(3).Caption = "-"
            txtOS(0).Text = ""
            txtOS(1).Text = ""
            txtOS(2).Text = ""
            Frame3.Visible = False
            txtOS(0).SetFocus
            MapaOs
            Frame9.Visible = True
        End If
    End Select
    Exit Sub
Err:
    Select Case Err.Number
    Case -2147467259
        txtOS(Index).Text = ""
        msgLabel "Falha na Conexão. Entre em contato com o Administrador da Rede", 2
        Timer2.Enabled = True
    End Select
End Sub

Private Function verificaSeApropria()
    verificaSeApropria = False
    Dim rsVerificaSeApropria As New ADODB.Recordset
    Dim SqlVerificaSeApropria As String

    'Ze louro e serginho nao passam por essa condição porque estao alocados diretamente no centro de custo principal
    SqlVerificaSeApropria = "Select * from tbApropriacao where substring(codreduzido,1,15) = '" & Mid$(lblOS(3), 1, 15) & "'"
    rsVerificaSeApropria.Open SqlVerificaSeApropria, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVerificaSeApropria.RecordCount > 0 Then
        verificaSeApropria = True
    End If
    rsVerificaSeApropria.Close
    Set rsVerificaSeApropria = Nothing
End Function

Private Function VerificaPermissao()
    VerificaPermissao = False
    Dim rsVerificaPermissao As New ADODB.Recordset
    Dim SqlVerificaPermissao As String
    
    SqlVerificaPermissao = "Select * from tbAutCCusto where chapa = '" & txtOS(0) & "' and idcc = '1'"
    rsVerificaPermissao.Open SqlVerificaPermissao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVerificaPermissao.RecordCount > 0 Then
        VerificaPermissao = False
        rsVerificaPermissao.Close
        Set rsVerificaPermissao = Nothing
        Exit Function
    End If
    rsVerificaPermissao.Close
    Set rsVerificaPermissao = Nothing
    
    SqlVerificaPermissao = "Select * from tbAutCCusto where chapa = '" & txtOS(0) & "'"
    rsVerificaPermissao.Open SqlVerificaPermissao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVerificaPermissao.RecordCount > 0 Then
        VerificaPermissao = True
    End If
    rsVerificaPermissao.Close
    Set rsVerificaPermissao = Nothing
    Exit Function
End Function

Private Function validaCracha()
    Dim rsAchaCC As New ADODB.Recordset
    Dim SqlAchaCC As String
    
    Dim rsValidaCracha As New ADODB.Recordset
    Dim SqlvalidaCracha As String
    
    validaCracha = False
    SqlAchaCC = "Select a.codigobarra,b.idos,b.idoperacao,b.idcc from tbOsMov as a inner join tbmpitens as b on a.codigobarra = b.codigobarra " & _
                "where a.chapa = '" & txtOS(0) & "' and datasai is null"
    rsAchaCC.Open SqlAchaCC, cnBanco, adOpenKeyset, adLockReadOnly
    
    vCBarraGeral = rsAchaCC.Fields(0)
    
    SqlvalidaCracha = "select * from tbautCCusto where chapa = '" & txtOS(0).Text & "' and idcc = '" & rsAchaCC.Fields(3) & "'"
    rsValidaCracha.Open SqlvalidaCracha, cnBanco, adOpenKeyset, adLockReadOnly
    If rsValidaCracha.RecordCount > 0 Then
        validaCracha = True
    End If
    rsAchaCC.Close
    Set rsAchaCC = Nothing
    
    rsValidaCracha.Close
    Set rsValidaCracha = Nothing
    
    Exit Function
End Function

Private Sub MapaOs()
    Dim rsMapaOS As New ADODB.Recordset
    Dim SqlMapaOS As String
    Dim numOS As Integer, contaLinha As Integer
    Dim ItemLst As ListItem
    Dim lvLoop As ListView
    LimpaLV ListView4
    LimpaLV ListView5
    LimpaLV ListView6
    LimpaLV ListView7
    
    SqlMapaOS = "select a.idprogramacao as os, d.DESCRICAO as [Onde Esta],b.dataent dataent,b.horaent horaent from tbMPItens as a inner join tbOsMov as b " & _
               "on a.codigobarra = b.codigobarra inner join CORPORERM.dbo.PFUNC as c on b.chapa COLLATE SQL_Latin1_General_CP1_CI_AS = c.CHAPA inner join CORPORERM.dbo.PSECAO as d " & _
               "on C.CODSECAO = D.CODIGO where a.status = 2 ORDER BY A.idprogramacao, B.dataent desc,b.horaent desc"
    rsMapaOS.Open SqlMapaOS, cnBanco, adOpenKeyset, adLockReadOnly
    If rsMapaOS.RecordCount <> 0 Then
        vnumos = rsMapaOS.Fields(0)
        Set ItemLst = ListView4.ListItems.Add(, , Format(rsMapaOS.Fields(0), "000000000"))
        ItemLst.SubItems(1) = "" & rsMapaOS.Fields(1)
        rsMapaOS.MoveNext
    Else
        Exit Sub
    End If
    contaLinha = 1
    Set lvLoop = ListView4
    
    While Not rsMapaOS.EOF
        If rsMapaOS.Fields(0) <> vnumos Then
            Set ItemLst = lvLoop.ListItems.Add(, , Format(rsMapaOS.Fields(0), "000000000"))
            ItemLst.SubItems(1) = "" & rsMapaOS.Fields(1)
            vnumos = rsMapaOS.Fields(0)
            contaLinha = contaLinha + 1
        End If
        rsMapaOS.MoveNext
        If contaLinha >= 15 And contaLinha <= 29 Then
            Set lvLoop = ListView5
        ElseIf contaLinha >= 30 And contaLinha <= 44 Then
            Set lvLoop = ListView6
        ElseIf contaLinha >= 45 And contaLinha <= 60 Then
            Set lvLoop = ListView7
        End If
    Wend
    Me.ListView4.Sorted = True
    Me.ListView4.SortKey = 0
    Me.ListView4.SortOrder = lvwAscending
    Me.ListView5.Sorted = True
    Me.ListView5.SortKey = 0
    Me.ListView5.SortOrder = lvwAscending
    Me.ListView6.Sorted = True
    Me.ListView6.SortKey = 0
    Me.ListView6.SortOrder = lvwAscending
    Me.ListView7.Sorted = True
    Me.ListView7.SortKey = 0
    Me.ListView7.SortOrder = lvwAscending
    rsMapaOS.Close
    Set rsMapaOS = Nothing
End Sub

Private Sub txtOS_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtOS_LostFocus(Index As Integer)
'    Select Case Index
'    Case 0
'        If vSetFocus = 2 Then
'            txtOS(0).SetFocus
'        End If
'    Case 1
'        If vSetFocus = 2 Then
'            txtOS(1).SetFocus
'        Else
'            txtOS(0).SetFocus
'        End If
'    Case 2
'        If vSetFocus = 2 Then
'            txtOS(2).SetFocus
'        Else
'            txtOS(0).SetFocus
'        End If
'    End Select
End Sub

Private Function compoeControles()
    compoeControles = False
    Dim rsCompoe As New ADODB.Recordset
    Dim SqlCompoe As String
    
    SqlCompoe = ""
    SqlCompoe = SqlCompoe & "SELECT " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CHAPA, " & vbCrLf
    SqlCompoe = SqlCompoe & " B.NOME, " & vbCrLf
    SqlCompoe = SqlCompoe & " C.NOME AS FUNCAO, " & vbCrLf
    SqlCompoe = SqlCompoe & " D.DESCRICAO AS SETOR, " & vbCrLf
    SqlCompoe = SqlCompoe & " F.NOME " & vbCrLf
    SqlCompoe = SqlCompoe & "FROM CORPORERM.DBO.PFUNC AS A " & vbCrLf
    SqlCompoe = SqlCompoe & "INNER JOIN CORPORERM.DBO.PPESSOA AS B ON " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODSITUACAO IN('A','F','P','Z','D') AND " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODPESSOA = B.CODIGO AND " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODSITUACAO <> 'D' " & vbCrLf
    SqlCompoe = SqlCompoe & "  " & vbCrLf
    SqlCompoe = SqlCompoe & " OR " & vbCrLf
    SqlCompoe = SqlCompoe & "  " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODSITUACAO IN('A','F','P','Z','D') AND " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODPESSOA = B.CODIGO AND A.CODSITUACAO = 'D' AND " & vbCrLf
    SqlCompoe = SqlCompoe & " GETDATE ( )<A.DTDESLIGAMENTO+1 " & vbCrLf
    SqlCompoe = SqlCompoe & "INNER JOIN CORPORERM.DBO.PFUNCAO AS C ON " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODFUNCAO = C.CODIGO " & vbCrLf
    SqlCompoe = SqlCompoe & "INNER JOIN CORPORERM.DBO.PSECAO AS D ON " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODSECAO = D.CODIGO AND " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODCOLIGADA = D.CODCOLIGADA " & vbCrLf
    SqlCompoe = SqlCompoe & "INNER JOIN CORPORERM.DBO.PFRATEIOFIXO AS E ON " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CHAPA = E.CHAPA " & vbCrLf
    SqlCompoe = SqlCompoe & "INNER JOIN CORPORERM.DBO.GCCUSTO AS F ON " & vbCrLf
    SqlCompoe = SqlCompoe & " E.CODCCUSTO = F.CODCCUSTO AND " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CODCOLIGADA = F.CODCOLIGADA " & vbCrLf
    SqlCompoe = SqlCompoe & "WHERE " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CHAPA = '" & Format(txtOS(0).Text, "00000") & "'" & vbCrLf
    SqlCompoe = SqlCompoe & " " & vbCrLf
    SqlCompoe = SqlCompoe & "UNION ALL " & vbCrLf
    SqlCompoe = SqlCompoe & " " & vbCrLf
    SqlCompoe = SqlCompoe & "SELECT " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CHAPA COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    SqlCompoe = SqlCompoe & " A.NOME COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    SqlCompoe = SqlCompoe & " A.FUNCAO COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    SqlCompoe = SqlCompoe & " A.SETOR COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI, " & vbCrLf
    SqlCompoe = SqlCompoe & " A.IDCC COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI + ' - ' + A.NMCC COLLATE SQL_LATIN1_GENERAL_CP1_CI_AI " & vbCrLf
    SqlCompoe = SqlCompoe & "FROM TBTERCEIRIZADOS AS A " & vbCrLf
    SqlCompoe = SqlCompoe & "LEFT JOIN TBAUTFECHAOS AS B ON " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CHAPA = B.CHAPA " & vbCrLf
    SqlCompoe = SqlCompoe & "   " & vbCrLf
    SqlCompoe = SqlCompoe & "WHERE " & vbCrLf
    SqlCompoe = SqlCompoe & " A.CHAPA = '" & Format(txtOS(0).Text, "00000") & "' AND " & vbCrLf
    SqlCompoe = SqlCompoe & " A.ATIVO = 'S'"
    
    rsCompoe.Open SqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
        msgLabel "", 1
        txtOS(0).Text = Format(txtOS(0).Text, "00000")
        lblOS(0).Caption = rsCompoe.Fields(1) 'Nome
        lblOS(1).Caption = rsCompoe.Fields(3) 'Setor
        lblOS(2).Caption = rsCompoe.Fields(2) 'Função
        lblOS(3).Caption = rsCompoe.Fields(4) 'Centro de Custo
        compoeControles = True
    Else
        msgLabel "Chapa não identificada no sistema", 2
        txtOS(0).Text = ""
        lblOS(0).Caption = "-"
        lblOS(1).Caption = "-"
        lblOS(2).Caption = "-"
        lblOS(3).Caption = "-"
        txtOS(0).SetFocus
    End If
    rsCompoe.Close
    Set rsCompoe = Nothing
End Function

Private Sub procuraSaida()
    Dim rsProcuraSaida As New ADODB.Recordset
    Dim SqlProcuraSaida As String
    SqlProcuraSaida = "select * from tbautccusto as a inner join tbParadas as b on b.codigo = a.idcc where chapa = '" & txtOS(0).Text & "'"
    rsProcuraSaida.Open SqlProcuraSaida, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProcuraSaida.RecordCount > 0 Then
        If rsProcuraSaida.Fields(3) = "Fechamento" Then
            frmConfirmaChapa.Show 1
            If vNomeGlobal = lblOS(0).Caption And vMSGGlobal = "OK" Then
                End
            Else
                msgLabel "Registro não identificado para EXECUTAR essa operação ", 2
                txtOS(0).Text = ""
            End If
        End If
    End If
    rsProcuraSaida.Close
    Set rsProcuraSaida = Nothing
End Sub
Private Sub achaEntrada()
    'sub para encontrar registros de entrada e saida na tabela tbOSMov
    Dim rsachaEntrada As New ADODB.Recordset
    Dim SqlachaEntrada As String
    slqachaEntrada = "select * from tbOsMov as a where a.chapa = '" & txtOS(0).Text & "' and a.datasai is null"
    rsachaEntrada.Open slqachaEntrada, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsachaEntrada.EOF And vVerificaPermissao = 0 Then
        'Parada
        Frame3.Visible = True
    Else
        'OS
        Frame2.Visible = True
    End If
    rsachaEntrada.Close
    Set rsachaEntrada = Nothing
End Sub

Private Function achaParada()
    'sub para encontrar registros de entrada e saida na tabela tbOSMov
    On Error Resume Next
    achaParada = False
    Dim rsachaParada As New ADODB.Recordset
    Dim SqlachaParada As String
    SqlachaParada = "select * from tbParadas as a where a.codigo = '" & txtOS(2).Text & "'"
    rsachaParada.Open SqlachaParada, cnBanco, adOpenKeyset, adLockReadOnly
    If rsachaParada.State <> adStateOpen Then
        Err.Number = -2147467259
        achaParada = True
        Exit Function
    End If
    If rsachaParada.EOF Then
        msgLabel "Código de PARADA não encontrado", 2
        txtOS(2).Text = ""
    Else
        lblTipoOS = rsachaParada.Fields(1)
        achaParada = True
    End If
    rsachaParada.Close
    Set rsachaParada = Nothing
End Function

Private Sub achaOS()
    'sub para encontrar, verificar status e registrar a OS se estiver liberada
    Dim rsAchaOS As New ADODB.Recordset
    Dim SqlAchaOs As String
    
    Dim rsLiberaAprop As New ADODB.Recordset
    Dim SqlLiberaAprop As String
    'SqlLiberaAprop = "select * from tbmpitens as a where a.codigobarra = '" & txtOS(1).Text & "'"
    SqlLiberaAprop = "select a.*,c.fce,d.status from tbmpitens as a left join tbMP as b on a.idprogramacao = b.idprogramacao left join tbProjetos as c on b.codprojeto = c.codprojeto left join tbFCE as d on c.fce = d.fce where a.codigobarra = '" & txtOS(1).Text & "'"
    rsLiberaAprop.Open SqlLiberaAprop, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsLiberaAprop.EOF Then
        If IsNull(rsLiberaAprop.Fields(5)) Then
            msgLabel "OS não possui PROGRAMAÇÃO. Favor informar ao setor de planejamento " & vCBarraGeral, 2
            rsLiberaAprop.Close
            Set rsLiberaAprop = Nothing
            txtOS(1).Text = ""
            txtOS(1).SetFocus
            Exit Sub
        ElseIf rsLiberaAprop.Fields(19) = 2 Then
            msgLabel "OS PARALISADA. Solicite informações com o responsável pelo setor" & vCBarraGeral, 2
            rsLiberaAprop.Close
            Set rsLiberaAprop = Nothing
            txtOS(1).Text = ""
            txtOS(1).SetFocus
            Exit Sub
        End If
    End If
    
    rsLiberaAprop.Close
    Set rsLiberaAprop = Nothing
    
    SqlAchaOs = "select * from tbos as a inner join tbositens as b on a.idos = b.idos where b.codigobarra = '" & txtOS(1).Text & "'"
    rsAchaOS.Open SqlAchaOs, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsAchaOS.EOF Then
        'Encontrou a OS
        If rsAchaOS.Fields(15) = 3 Then
            'OS esta fechada
            msgLabel "Essa OPERAÇÃO encontra-se FECHADA. Não pode ser registrada.", 2
            txtOS(1).Text = ""
            txtOS(1).SetFocus
        Else
            'SE O COLABORADOR TIVER PERMISSÃO APENAS PARA FECHAR OS
            'ENTRA NO BLOCO DE CONDIÇÕES ABAIXO
            '=================================================================
            If vVerificaPermissao = 1 Then
                frmConfirmaChapa.Show 1
                If vNomeGlobal = lblOS(0).Caption And vMSGGlobal = "OK" Then
                    fechaOS
                Else
                    msgLabel "Colaborador: " & vNomeGlobal & " não tem permissão para FECHAR a operação: " & vCBarraGeral, 2
                End If
                lblOS(0).Caption = "-"
                lblOS(1).Caption = "-"
                lblOS(2).Caption = "-"
                lblOS(3).Caption = "-"
                txtOS(0).Text = ""
                txtOS(1).Text = ""
                txtOS(2).Text = ""
                Frame3.Visible = False
                txtOS(0).SetFocus
                MapaOs
                Frame9.Visible = True
                Exit Sub
            End If
            '=================================================================
            
            'OS esta em aberto
            msgLabel lblOS(0).Caption & " registrou ENTRADA na OS nº: " & Format(rsAchaOS.Fields(0), "000000000") & ", OP. nº: " & rsAchaOS.Fields(17) & "  - C.C.: " & rsAchaOS.Fields(13), 1
            'msgLabel "OS nº: " & txtOS(1).Text & " - C.C.: " & rsachaOS.Fields(8) & " - Registrada para: " & lblOS(0).Caption, 1
            gravaOS
            iniciaRetrabalho txtOS(1).Text
            lblOS(0).Caption = "-"
            lblOS(1).Caption = "-"
            lblOS(2).Caption = "-"
            lblOS(3).Caption = "-"
            txtOS(0).Text = ""
            txtOS(1).Text = ""
            txtOS(2).Text = ""
            Frame2.Visible = False
            txtOS(0).SetFocus
            MapaOs
            Frame9.Visible = True
        End If
    Else
        'Não encontrou a OS
        msgLabel "OS não encontrada", 2
        txtOS(1).Text = ""
    End If
    rsAchaOS.Close
    Set rsAchaOS = Nothing
    Exit Sub
End Sub

Private Sub gravaOS()
    Dim rsGravaOS As New ADODB.Recordset
    Dim SqlGravaOS As String
    Dim rsAlteraStatus As New ADODB.Recordset
    Dim SqlAlteraStatus As String
    Dim rsAlteraStatusOS As New ADODB.Recordset
    Dim SqlAlteraStatusOS As String
    Dim rsAlteraStatusITENS As New ADODB.Recordset
    Dim SqlAlteraStatusITENS As String
    Dim rsAlteraStatusMP As New ADODB.Recordset
    Dim SqlAlteraStatusMP As String
    Dim rsLimpaStatusMP As New ADODB.Recordset
    Dim SqlLimpaStatusMP As String
    Dim vIdOs As Integer, vProg As Integer
    Dim vCodigoBarra As String, vCC As String
    
    Dim rsPegaDataHoraServer As New ADODB.Recordset
    Dim SqlPegaDataHoraServer As String
    Dim vDataServer As String, vHoraServer As String
    SqlPegaDataHoraServer = "select CONVERT (VARCHAR, CURRENT_TIMESTAMP,103) as dataServidor, CONVERT (VARCHAR, CURRENT_TIMESTAMP,108) as horaServidor"
    rsPegaDataHoraServer.Open SqlPegaDataHoraServer, cnBanco, adOpenKeyset, adLockReadOnly
    vDataServer = rsPegaDataHoraServer.Fields(0)
    vHoraServer = rsPegaDataHoraServer.Fields(1)
    rsPegaDataHoraServer.Close
    Set rsPegaDataHoraServer = Nothing
    
    
    
    'SE NO CAMPO RESERVADO PARA O NUMERO DA OS FOR INSERIDO O CÓDIGO 1000
    'IRÁ SER REGISTRADO UMA PARADA NA ENTRADA DO SERVIÇO
    If txtOS(1).Text <> "1000" Then
        SqlGravaOS = "Insert into tbOsMov(chapa,codigobarra,dataent,horaent) Values('" & txtOS(0).Text & "','" & txtOS(1).Text & "','" & Format(vDataServer, "YYYY-MM-DD") & "','" & vHoraServer & "')"
    Else
        SqlGravaOS = "Insert into tbOsMov(chapa,codigobarra,dataent,horaent) Values('" & txtOS(0).Text & "','" & txtOS(2).Text & "','" & Format(vDataServer, "YYYY-MM-DD") & "','" & vHoraServer & "')"
    End If
    rsGravaOS.Open SqlGravaOS, cnBanco
    
    SqlAlteraStatus = "select a.codigobarra,c.idos,b.idcc,b.status,b.idprogramacao from tbosmov as a inner join tbositens as b on a.codigobarra = b.codigobarra inner join tbos as c " & _
                      "on b.idos = c.idos where a.codigobarra = '" & txtOS(1).Text & "' and b.status = 1 group by a.codigobarra,c.idos,b.idcc,b.status,b.idprogramacao"
    rsAlteraStatus.Open SqlAlteraStatus, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAlteraStatus.RecordCount > 0 Then
        'ALTERA O STATUS DA OS PARA: 2 - ANDAMENTO
        vCodigoBarra = rsAlteraStatus.Fields(0)
        vIdOs = rsAlteraStatus.Fields(1)
        vCC = rsAlteraStatus.Fields(2)
        vProg = rsAlteraStatus.Fields(4) 'nº da programacao
        
        SqlAlteraStatusOS = "Update tbOs set status = 2 where idos = '" & vIdOs & "'"
        rsAlteraStatusOS.Open SqlAlteraStatusOS, cnBanco
        
        SqlAlteraStatusITENS = "Update tbOsItens set status = 2 where codigobarra = '" & vCodigoBarra & "'"
        rsAlteraStatusITENS.Open SqlAlteraStatusITENS, cnBanco
        
        SqlAlteraStatusMP = "Update tbMPItens set status = 2 where codigobarra = '" & vCodigoBarra & "'"
        rsAlteraStatusMP.Open SqlAlteraStatusMP, cnBanco
        
        SqlAlteraStatusMP = "Update tbMPItens set databaixa = GetDate() where codigobarra = '" & vCodigoBarra & "' and status = 3 and databaixa is null"
        rsAlteraStatusMP.Open SqlAlteraStatusMP, cnBanco
        
        SqlLimpaStatusMP = "Update tbMP set status = NULL where idprogramacao = '" & vProg & "'"
        rsLimpaStatusMP.Open SqlLimpaStatusMP, cnBanco
        
    End If
    rsAlteraStatus.Close
    Set rsAlteraStatus = Nothing
End Sub

Private Sub gravaParada()
    Dim rsGravaOS As New ADODB.Recordset
    Dim SqlGravaOS As String
    
    Dim rsPegaDataHoraServer As New ADODB.Recordset
    Dim SqlPegaDataHoraServer As String
    Dim vDataServer As String, vHoraServer As String
    SqlPegaDataHoraServer = "select CONVERT (VARCHAR, CURRENT_TIMESTAMP,103) as dataServidor, CONVERT (VARCHAR, CURRENT_TIMESTAMP,108) as horaServidor"
    rsPegaDataHoraServer.Open SqlPegaDataHoraServer, cnBanco, adOpenKeyset, adLockReadOnly
    vDataServer = rsPegaDataHoraServer.Fields(0)
    vHoraServer = rsPegaDataHoraServer.Fields(1)
    rsPegaDataHoraServer.Close
    Set rsPegaDataHoraServer = Nothing
    
    SqlGravaOS = "Update tbOsMov set datasai = '" & Format(vDataServer, "YYYY-MM-DD") & "',horasai = '" & vHoraServer & "', idparada ='" & txtOS(2).Text & "' where chapa = '" & txtOS(0).Text & "' and datasai is null"
    rsGravaOS.Open SqlGravaOS, cnBanco
End Sub

Private Sub gravaParada2() 'Fecha todos os movimentos da OPERACAO
    Dim rsGravaOS As New ADODB.Recordset
    Dim SqlGravaOS As String
    
    Dim rsPegaDataHoraServer As New ADODB.Recordset
    Dim SqlPegaDataHoraServer As String
    Dim vDataServer As String, vHoraServer As String
    SqlPegaDataHoraServer = "select CONVERT (VARCHAR, CURRENT_TIMESTAMP,103) as dataServidor, CONVERT (VARCHAR, CURRENT_TIMESTAMP,108) as horaServidor"
    rsPegaDataHoraServer.Open SqlPegaDataHoraServer, cnBanco, adOpenKeyset, adLockReadOnly
    vDataServer = rsPegaDataHoraServer.Fields(0)
    vHoraServer = rsPegaDataHoraServer.Fields(1)
    rsPegaDataHoraServer.Close
    Set rsPegaDataHoraServer = Nothing
    
    SqlGravaOS = "Update tbOsMov set datasai = '" & Format(vDataServer, "YYYY-MM-DD") & "',horasai = '" & vHoraServer & "', idparada ='" & txtOS(2).Text & "' where codigobarra = '" & vCBarraGeral & "' and datasai is null"
    rsGravaOS.Open SqlGravaOS, cnBanco
    msgLabel "OPERAÇÃO nº: " & vCBarraGeral & " fechada por: " & lblOS(0).Caption, 1
End Sub

Private Sub registraFechamento()
    'REGISTRA APROPRIAÇÃO DE FECHAMENTO DO ENCARREGADO/CONTRAMESTRE
    Dim rsRegistraFechamento As New ADODB.Recordset
    Dim SqlRegistraFechamento As String
    Dim vParadaFechar As String
    vParadaFechar = "9020"
    SqlRegistraFechamento = "Insert into tbOsMov(chapa,codigobarra,dataent,horaent,datasai,horasai,idparada) Values('" & txtOS(0).Text & "','" & txtOS(1).Text & "','" & Format(Date, "YYYY-MM-DD") & "','" & Time & "','" & Format(Date, "YYYY-MM-DD") & "','" & Time & "','" & vParadaFechar & "')"
    rsRegistraFechamento.Open SqlRegistraFechamento, cnBanco
End Sub


Private Sub fechaOS()
    'msgLabel "Inicio da rotina de FECHAMENTO da OS", 1
    
    Dim rsAchaOS As New ADODB.Recordset
    Dim SqlAchaOs As String
    Dim rsMontaDados As New ADODB.Recordset
    Dim SqlMontaDados As String
    
    Dim rsAlteraStatusOS As New ADODB.Recordset
    Dim SqlAlteraStatusOS As String
    Dim rsAlteraStatusITENS As New ADODB.Recordset
    Dim SqlAlteraStatusITENS As String
    Dim rsAlteraStatusMP As New ADODB.Recordset
    Dim SqlAlteraStatusMP As String
    Dim rsLimpaStatusMP As New ADODB.Recordset
    Dim SqlLimpaStatusMP As String
    
    Dim rsVerificaDados As New ADODB.Recordset
    Dim SqlVerificaDados As String
    
    Dim vIdOs As Integer, vOperacao As Integer, vLiberaInsp As Integer, vProgramacao As Integer
    Dim vCodigoBarra As String, vCC As String
    Dim vEnviaEmail As Integer
    Dim vQuemEstaLiberando As String
    
    'ABAIXO QUERY PARA LOCALIZAR OS DADOS
    
    If vVerificaPermissao = 0 Then
        SqlAchaOs = "Select a.codigobarra,b.idos,b.idoperacao,b.idcc,b.idprogramacao from tbOsMov as a inner join tbmpitens as b " & _
                    "on a.codigobarra = b.codigobarra where a.chapa = '" & txtOS(0).Text & "' and datasai is null"
    Else
        SqlAchaOs = "select a.codigobarra,a.idos,a.idoperacao,a.idcc,a.idprogramacao from tbMPItens as a where a.codigobarra = '" & txtOS(1).Text & "'"
    End If
    
    rsAchaOS.Open SqlAchaOs, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsAchaOS.RecordCount > 0 Then
        vCodigoBarra = rsAchaOS.Fields(0)
        vCBarraGeral = rsAchaOS.Fields(0)
        vIdOs = rsAchaOS.Fields(1)
        vOperacao = rsAchaOS.Fields(2)
        vCC = rsAchaOS.Fields(3)
        vProgramacao = rsAchaOS.Fields(4)
        vLiberaInsp = 0
        'Abaixo query do loop
        'SqlMontaDados = "select a.idprogramacao,a.idsequencia,a.idcc,a.grupo,a.idos,a.idoperacao,a.codigobarra,a.status,b.formula from tbMPItens as a inner join tbFormula as b on a.idcc = b.codreduzido and b.idform = 1 " & _
        '                "where a.idprogramacao = '" & vProgramacao & "' and a.idos ='" & vIdOs & "' and a.idoperacao ='" & vOperacao & "' and a.status <> 3 order by a.idoperacao"
'-------- Linhas em TESTE ABAIXO -----------------------
        '1º) Entra se for centro de custo da qualidade
        If Mid$(lblOS(3).Caption, 1, 9) = "7000.7103" Then
            SqlMontaDados = "select a.idprogramacao,a.idsequencia,a.idcc,a.grupo,a.idos,a.idoperacao,a.codigobarra,a.status,b.formula from tbMPItens as a inner join tbFormula as b on a.idcc = b.codreduzido and b.idform = 1 " & _
                            "where a.idprogramacao = '" & vProgramacao & "' and a.idos ='" & vIdOs & "' and a.idoperacao <='" & vOperacao & "' and a.status <> 3 order by a.idoperacao"
        '2º) Entra se NÃO for centro de custo da qualidade
        Else
            SqlMontaDados = "select a.idprogramacao,a.idsequencia,a.idcc,a.grupo,a.idos,a.idoperacao,a.codigobarra,a.status,b.formula from tbMPItens as a inner join tbFormula as b on a.idcc = b.codreduzido and b.idform = 1 " & _
                            "where a.idprogramacao = '" & vProgramacao & "' and a.idos ='" & vIdOs & "' and a.idoperacao ='" & vOperacao & "' and a.status <> 3 order by a.idoperacao"
        End If
'-------- Linhas em TESTE ACIMA -----------------------
        
        
        rsMontaDados.Open SqlMontaDados, cnBanco, adOpenKeyset, adLockReadOnly
        If rsMontaDados.RecordCount = 0 Then
            msgLabel "Operação encontra-se FECHADA", 2
            rsMontaDados.Close
            Set rsMontaDados = Nothing
            rsAchaOS.Close
            Set rsAchaOS = Nothing
            Exit Sub
        End If
        rsMontaDados.MoveLast
        If rsMontaDados.Fields(8) = "LD" Then 'VERIFICA SE O ULTIMO REGISTRO DA CONSULTA É LIBERAÇÃO DIRETA
            vLiberaInsp = rsMontaDados.Fields(5) 'INFORMA QUAL OPERACAO DEVERA SER LIBERADA
            vQuemEstaLiberando = rsMontaDados.Fields(8)
        End If
        
        '-------------------------------------------------------
        'ABAIXO VERIFICA SE O CC FECHADO É ACABAMENTO. SE FOR:
        'vEnviaEmail = 1 - Marca para enviar email
        'vEnviaEmail = 0 - Marca para não enviar email
        If Mid$(rsMontaDados.Fields(2), 1, 9) = "3000.3105" Then
            vEnviaEmail = 1
        Else
            vEnviaEmail = 0
        End If
        '--------------------------------------------------------
                
        rsMontaDados.MoveFirst
        While Not rsMontaDados.EOF
            If rsMontaDados.Fields(8) <> "LD" And vLiberaInsp = 0 Or rsMontaDados.Fields(8) = "LD" And rsMontaDados.Fields(5) = vLiberaInsp Or vQuemEstaLiberando = "LD" Then
                SqlAlteraStatusITENS = "Update tbOsItens set status = 3 where codigobarra = '" & rsMontaDados.Fields(6) & "'"
                rsAlteraStatusITENS.Open SqlAlteraStatusITENS, cnBanco
        
                SqlAlteraStatusMP = "Update tbMPItens set status = 3 where codigobarra = '" & rsMontaDados.Fields(6) & "'"
                rsAlteraStatusMP.Open SqlAlteraStatusMP, cnBanco
                
                SqlAlteraStatusMP = "Update tbMPItens set databaixa = GetDate() where codigobarra = '" & rsMontaDados.Fields(6) & "' and status = 3 and databaixa is null"
                rsAlteraStatusMP.Open SqlAlteraStatusMP, cnBanco
                
                'Em Teste
                vCBarraGeral = rsMontaDados.Fields(6)
                gravaParada2
                'Em Teste
                
                'A rotina abaixo é referente ao CC de Inspeção de Qualidade
                'Ela verifica se a OS que está sendo fechada é uma OS de RETRABALHO
                'Se for ela irá entrar em duas tabelas (tbRNC/tbComunicacaoDesvio) e alterar o status para 10
                fechaRetrabalho vCBarraGeral
            Else
                msgLabel "Operação não pode ser FECHADA devido a pendência de INSPEÇÃO", 2
                rsMontaDados.Close
                Set rsMontaDados = Nothing
                rsAchaOS.Close
                Set rsAchaOS = Nothing
                Exit Sub
            End If
            rsMontaDados.MoveNext
        Wend
        
        SqlLimpaStatusMP = "Update tbMP set status = NULL where idprogramacao = '" & vProgramacao & "'"
        rsLimpaStatusMP.Open SqlLimpaStatusMP, cnBanco
        
        
        'REGISTRA APROPRIAÇÃO DE FECHAMENTO DO ENCARREGADO
        registraFechamento
        
        SqlVerificaDados = "select * from tbMPItens as a where a.idos ='" & vIdOs & "' and a.status <> 3 "
        rsVerificaDados.Open SqlVerificaDados, cnBanco, adOpenKeyset, adLockReadOnly
        
        If rsVerificaDados.RecordCount = 0 Then
            'Fechamento geral da OS
            SqlAlteraStatusOS = "Update tbOs set status = 3 where idos = '" & vIdOs & "'"
            rsAlteraStatusOS.Open SqlAlteraStatusOS, cnBanco
        End If
        
'-------Teste
        If Mid$(lblOS(3).Caption, 1, 9) = "7000.7103" Then
            SqlAlteraStatusMP = "Update tbMPItens set status = 3 where idos = '" & vIdOs & "'"
            rsAlteraStatusMP.Open SqlAlteraStatusMP, cnBanco
            enviaEmailLogistica vIdOs, vCodigoBarra, vOperacao, vCC
        End If
'-------Teste

        
        'Aki chamar rotina de envio de email
        If vEnviaEmail = 1 Then
            
            Dim rsAchaOPQualidade As New ADODB.Recordset
            Dim SqlAchaOPQualidade As String
            Dim vCodigoBarraQuality As String
            SqlAchaOPQualidade = "select a.codigobarra from tbMPItens as a where a.idos = '" & vIdOs & "' and SUBSTRING(a.idcc,1,4) = '7000'"
            rsAchaOPQualidade.Open SqlAchaOPQualidade, cnBanco, adOpenKeyset, adLockReadOnly
            vCodigoBarraQuality = rsAchaOPQualidade.Fields(0)
            
            If vSMTP <> "" Then enviaEmailQualidade vIdOs, vCodigoBarra, vOperacao, vCC, vCodigoBarraQuality
            rsAchaOPQualidade.Close
            Set rsAchaOPQualidade = Nothing
            
        End If
        rsMontaDados.Close
        Set rsMontaDados = Nothing
        rsAchaOS.Close
        Set rsAchaOS = Nothing
    End If
End Sub

Private Sub enviaEmailQualidade(vEIdOs As Integer, vECodigoBarra As String, vEOperacao As Integer, vECC As String, vECodigoBarraQuality As String)
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
'On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    RestauraEmailEnvio "SI", sEmailSI
    
    vDesenhosGlobal = ""
    
    achaFCEDesenho vEIdOs, vEOperacao
    
    'vSMTP = "smtp.viga.ind.br"
    'vUsuEmail = "viga@viga.ind.br"
    'vSenhaEmail = "Xbkwolpb7rpd0td"
    
    carregaDadosEmail
    
    'vSMTP = "mail.viga.ind.br"
    'vUsuEmail = "taos@viga.ind.br"
    'vSenhaEmail = "taos2017@"
    
    vDecisao = "Aprovado"
    vCorDecisao = "#CD2626"

    With Camp
        
        
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
        
        
        
'        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
'        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
'        .Item(cdoSMTPServerPort) = "465"  'Porta
'        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
'        .Item(cdoSMTPAuthenticate) = 1
'        .Item(cdoSMTUseSSL).Value = False
'        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
'        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
'        .Update
        
        
'        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
'        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
'        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
'        .Item(cdoSMTPAuthenticate) = 1
'        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
'        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
'        .Update
    End With

    With Msg
        Set .Configuration = Cof
      
'       .To = "viga@viga.ind.br;qualidade@viga.ind.br;producao@viga.ind.br;planejamento4@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .To = sEmailSI
        .From = "viga@viga.ind.br"  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "Registro de Fechamento de Operação de Acabamento nº: " & vECodigoBarra
        
        .HTMLBody = "<html>" & _
        " <head>" & _
        " <meta http-equiv='Content-Type'" & _
        " content='text/html; charset=iso-8859-1'>" & _
        " <meta name='GENERATOR' content='Microsoft FrontPage Express 2.0'>" & _
        " <title>Assinatura</title>" & _
        " <STYLE type='text/css'>" & _
        " <!-- -->" & _
        " </STYLE></head>" & _
        " <body link='#0000FF' vlink='#800080'>" & _
        " <font face = 'Courier New' size = 5>" & _
        " <B><FONT STYLE='COLOR:#009ACD'> SOLICITAÇÃO DE INSPEÇÃO </FONT></B><BR><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> Foi realizado o fechamento da operação de Acabamento referente a: </FONT><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> </FONT><BR><BR><FONT STYLE='COLOR:#009ACD'> OS nº: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & vEIdOs & " </b><BR><FONT STYLE='COLOR:#009ACD'>Centro de Custo nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vECC & "</b></FONT><BR>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & "" & " </FONT><FONT STYLE='COLOR:#009ACD'> Operação nº: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & vEOperacao & " </b></FONT><BR><FONT STYLE='COLOR:#009ACD'>FCE nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vFCEGlobal & "</b></FONT><BR><FONT STYLE='COLOR:#009ACD'>Desenhos/Rev.nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vDesenhosGlobal & "</b></FONT><BR><BR>  <FONT STYLE='COLOR:#009ACD'> Controle de Qualidade (CB): </FONT><FONT STYLE='COLOR:#CD2626'><b> " & vECodigoBarraQuality & " </b></FONT>   <BR><BR></font>" & _
        " <table border='0' cellspacing='0' width='627'>" & _
        " <tr><td width='100%'><span class='txt'>" & _
        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
        " </td></tr></table></body>"
        .Send
    End With
    'mobjMsg.Abrir "Email enviado com suscesso!", Ok, informacao, "ZEUS"
    Exit Sub
errMail:
    MsgBox "Email não enviado para o usuário solicitante do PDO." & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SGCH"
    Exit Sub
End Sub

Private Sub enviaEmailLogistica(vEIdOs As Integer, vECodigoBarra As String, vEOperacao As Integer, vECC As String)
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    RestauraEmailEnvio "SRM", sEmailSRM
    
    vDesenhosGlobal = ""
    
    achaFCEDesenho vEIdOs, vEOperacao
    
    carregaDadosEmail
    
    'vSMTP = "mail.viga.ind.br"
    'vUsuEmail = "taos@viga.ind.br"
    'vSenhaEmail = "taos2017@"
    
    vDecisao = "Aprovado"
    vCorDecisao = "#CD2626"

    With Camp
        
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
        
        
'        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
'        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
'        .Item(cdoSMTPServerPort) = "465"  'Porta
'        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
'        .Item(cdoSMTPAuthenticate) = 1
'        .Item(cdoSMTUseSSL).Value = False
'        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
'        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
'        .Update
    End With

    With Msg
        Set .Configuration = Cof
      
'        .To = "viga@viga.ind.br;almoxarifado@viga.ind.br;planejamento3@viga.ind.br;planejamento4@viga.ind.br;qualidade@viga.ind.br" 'vEmailSolicitante  ' destinatario1@email.com.br;destinatario2@email.com.br ‘ destinatarios separados por ;
        .To = sEmailSRM
        .From = "viga@viga.ind.br"  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "Registro de Fechamento de Operação de QUALIDADE nº: " & vECodigoBarra
        
        .HTMLBody = "<html>" & _
        " <head>" & _
        " <meta http-equiv='Content-Type'" & _
        " content='text/html; charset=iso-8859-1'>" & _
        " <meta name='GENERATOR' content='Microsoft FrontPage Express 2.0'>" & _
        " <title>Assinatura</title>" & _
        " <STYLE type='text/css'>" & _
        " <!-- -->" & _
        " </STYLE></head>" & _
        " <body link='#0000FF' vlink='#800080'>" & _
        " <font face = 'Courier New' size = 5>" & _
        " <B><FONT STYLE='COLOR:#009ACD'> SOLICITAÇÃO DE RETIRADA DE MATERIAL </FONT></B><BR><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> Foi realizado o fechamento da operação de QUALIDADE referente a: </FONT><BR></font>" & _
        " <font face = 'Courier New' size = 2>" & _
        " <FONT STYLE='COLOR:#009ACD'> </FONT><BR><BR><FONT STYLE='COLOR:#009ACD'> OS nº: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & vEIdOs & " </b><BR><FONT STYLE='COLOR:#009ACD'>Centro de Custo nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vECC & "</b></FONT><BR>" & _
        " <FONT STYLE='COLOR:#009ACD'> " & "" & " </FONT><FONT STYLE='COLOR:#009ACD'> Operação nº: </FONT><FONT STYLE='COLOR:#CD2626'><b> " & vEOperacao & " </b></FONT><BR><FONT STYLE='COLOR:#009ACD'>FCE nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vFCEGlobal & "</b></FONT><BR><FONT STYLE='COLOR:#009ACD'>Desenhos/Rev.nº: <FONT STYLE='COLOR:" & vCorDecisao & "'><b>" & vDesenhosGlobal & "</b></FONT><BR><BR><BR><BR></font>" & _
        " <table border='0' cellspacing='0' width='627'>" & _
        " <tr><td width='100%'><span class='txt'>" & _
        " <FONT STYLE='COLOR:#009ACD'> AGUARDANDO A RETIRADA DO MATERIAL DA ÁREA </FONT><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><I><FONT STYLE='COLOR:#000080'> " & NomUsu & "</FONT></I></B><BR></font>" & _
        " <font face = 'Courier New' size = 2><B><FONT STYLE='COLOR:#008000'>Preserve o meio ambiente! Pense antes de imprimir</FONT></B></font>" & _
        " <font face = 'Webdings' size = 3><B><FONT STYLE='COLOR:#008000'> P </FONT></B></font>" & _
        " </td></tr></table></body>"
        .Send
    End With
    'mobjMsg.Abrir "Email enviado com suscesso!", Ok, informacao, "ZEUS"
    Exit Sub
errMail:
    MsgBox "Email não enviado para o almoxarifado." & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "SGCH"
    Exit Sub
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Cód.", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Nome Parada", ListView1.Width / 1.2
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Cód.", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Nome Parada", ListView2.Width / 1.2
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "Cód.", ListView3.Width / 6
    ListView3.ColumnHeaders.Add , , "Nome Parada", ListView3.Width / 1.2
    
    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "OS", ListView4.Width / 3.8
    ListView4.ColumnHeaders.Add , , "Último Registro", ListView4.Width / 1.4
    
    ListView5.ColumnHeaders.Clear
    ListView5.ColumnHeaders.Add , , "OS", ListView5.Width / 3.8
    ListView5.ColumnHeaders.Add , , "Último Registro", ListView5.Width / 1.4
    
    ListView6.ColumnHeaders.Clear
    ListView6.ColumnHeaders.Add , , "OS", ListView6.Width / 3.8
    ListView6.ColumnHeaders.Add , , "Último Registro", ListView6.Width / 1.4
    
    ListView7.ColumnHeaders.Clear
    ListView7.ColumnHeaders.Add , , "OS", ListView7.Width / 3.8
    ListView7.ColumnHeaders.Add , , "Último Registro", ListView7.Width / 1.4
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    ListView5.View = lvwReport 'Modo de Exibição do seu Listview
    ListView6.View = lvwReport 'Modo de Exibição do seu Listview
    ListView7.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

'Remover rotina abaixo apos teste
Private Sub Timer1_Timer()
    Segundos.Caption = Format(Segundos.Caption + 1, "00")
    If Segundos.Caption = "60" Then
        Minutos.Caption = Format(Minutos.Caption + 1, "00")
        Segundos.Caption = "00"
    End If
    If Minutos.Caption = "60" Then
        Horas.Caption = Format(Horas.Caption + 1, "00")
        Minutos.Caption = "00"
    End If
    Label21 = Time
End Sub

Private Sub Timer2_Timer()
    Conectar
    Timer2.Enabled = False
End Sub

Private Sub achaFCEDesenho(vOSEmail As Integer, vOperacaoEmail)
    Dim rsAchaOSEmail As New ADODB.Recordset
    Dim SqlAchaOSEmail As String
    
    SqlAchaOSEmail = "select a.desenhos,c.fce from tbMPItens as a inner join tbMP as b on a.idprogramacao = b.idprogramacao inner join tbProjetos as c on b.codprojeto = c.codprojeto WHERE idos = '" & vOSEmail & "' and idoperacao = '" & vOperacaoEmail & "'"
    rsAchaOSEmail.Open SqlAchaOSEmail, cnBanco, adOpenKeyset, adLockReadOnly
    vFCEGlobal = rsAchaOSEmail.Fields(1) 'Guarda o numero da FCE para ser enviado no EMAIL
    separaDadosDesenhos rsAchaOSEmail.Fields(0), rsAchaOSEmail.Fields(1)
End Sub


Private Sub separaDadosDesenhos(vTxtForm As String, vFCEEmail As Integer)
On Error Resume Next
    Dim rsAchaDesenho As New ADODB.Recordset
    Dim SqlAchaDesenho As String
    
    Dim rsCompoeDesenhos As New ADODB.Recordset
    Dim SqlCompoeDesenhos As String

    
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String, vFCE As String
    Dim vCodLM As String, vCodSeq As String
    
    SqlTransf = "Delete from tbTAOSEmail"
    rsTransf.Open SqlTransf, cnBanco
    
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            'Separa para localizar: codigo da LM e código da sequência da LM
            'Se a variável recebe tiver + de 5 caracteres significa que a sequencia da LM ultrapassou a 999 registros
            'O procedimento para esse caso é diferenciado, por isso utilizasse o IF abaixo
            If Len(RECEBE) = 5 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 3)
            ElseIf Len(RECEBE) = 6 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 4)
            End If
            
            SqlAchaDesenho = "select a.fce,a.codlm,a.codseq,b.desenho,b.revisao from tbItemLM as a inner join tbdesenhos as b on a.codigodes = b.iddesenho where a.fce = '" & vFCEEmail & "' and a.codlm = '" & vCodLM & "' and a.codseq = '" & vCodSeq & "'"
            rsAchaDesenho.Open SqlAchaDesenho, cnBanco, adOpenKeyset, adLockReadOnly
            While Not rsAchaDesenho.EOF
                SqlTransf = "Insert into tbTAOSEmail(fce,desenho,revisao) Values('" & Val(rsAchaDesenho.Fields(0)) & "','" & rsAchaDesenho.Fields(3) & "','" & rsAchaDesenho.Fields(4) & "')"
                rsTransf.Open SqlTransf, cnBanco
                
                rsAchaDesenho.MoveNext
            Wend
            rsAchaDesenho.Close
            'Set rsAchaDesenho = Nothing
            
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        'Separa para localizar: codigo da LM e código da sequência da LM
        'Se a variável recebe tiver + de 5 caracteres significa que a sequencia da LM ultrapassou a 999 registros
        'O procedimento para esse caso é diferenciado, por isso utilizasse o IF abaixo
        If Len(RECEBE) = 5 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 3)
        ElseIf Len(RECEBE) = 6 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 4)
        End If
        SqlAchaDesenho = "select a.fce,a.codlm,a.codseq,b.desenho,b.revisao from tbItemLM as a inner join tbdesenhos as b on a.codigodes = b.iddesenho where a.fce = '" & vFCEEmail & "' and a.codlm = '" & vCodLM & "' and a.codseq = '" & vCodSeq & "'"
        rsAchaDesenho.Open SqlAchaDesenho, cnBanco, adOpenKeyset, adLockReadOnly
        While Not rsAchaDesenho.EOF
            SqlTransf = "Insert into tbTAOSEmail(fce,desenho,revisao) Values('" & rsAchaDesenho.Fields(0) & "','" & rsAchaDesenho.Fields(3) & "','" & rsAchaDesenho.Fields(4) & "')"
            rsTransf.Open SqlTransf, cnBanco
            rsAchaDesenho.MoveNext
        Wend
        rsAchaDesenho.Close
        Set rsAchaDesenho = Nothing
        
        SqlCompoeDesenhos = "select * from tbTAOSEmail"
        rsCompoeDesenhos.Open SqlCompoeDesenhos, cnBanco, adOpenKeyset, adLockReadOnly
        While Not rsCompoeDesenhos.EOF
            If vDesenhosGlobal = "" Then
                vDesenhosGlobal = rsCompoeDesenhos.Fields(1) & " (" & rsCompoeDesenhos.Fields(2) & ")"
            Else
                vDesenhosGlobal = vDesenhosGlobal & "/" & rsCompoeDesenhos.Fields(1) & " (" & rsCompoeDesenhos.Fields(2) & ")"
            End If
            rsCompoeDesenhos.MoveNext
        Wend
        
        rsCompoeDesenhos.Close
        Set rsCompoeDesenhos = Nothing
    End If
End Sub

Private Sub iniciaRetrabalho(vCodBarra As String)
    Dim rsLocalizaRetrabalho As New ADODB.Recordset
    Dim SqlLocalizaRetrabalho As String
    SqlLocalizaRetrabalho = "select a.idprogramacao,b.idcd,b.idretrabalho from tbMPItens as a inner join tbRetrabalho as b on a.idprogramacao = b.idprogramacao where codigobarra = '" & vCodBarra & "'"
    rsLocalizaRetrabalho.Open SqlLocalizaRetrabalho, cnBanco, adOpenKeyset, adLockReadOnly
    If rsLocalizaRetrabalho.RecordCount > 0 Then
        Dim rsAlteraStatusRNC As New ADODB.Recordset
        Dim SqlAlteraStatusRNC As String
    
        Dim rsAlteraStatusCD As New ADODB.Recordset
        Dim SqlAlteraStatusCD As String
    
        SqlAlteraStatusRNC = "update tbRNC set status = 9 where idcd = '" & rsLocalizaRetrabalho.Fields(1) & "'"
        rsAlteraStatusRNC.Open SqlAlteraStatusRNC, cnBanco
        
        SqlAlteraStatusCD = "update tbComunicacaoDesvio set status = 9 where idcd = '" & rsLocalizaRetrabalho.Fields(1) & "'"
        rsAlteraStatusCD.Open SqlAlteraStatusCD, cnBanco
    End If
    rsLocalizaRetrabalho.Close
    Set rsLocalizaRetrabalho = Nothing
End Sub

Private Sub fechaRetrabalho(vCodBarra As String)
    Dim rsLocalizaRetrabalho As New ADODB.Recordset
    Dim SqlLocalizaRetrabalho As String
    SqlLocalizaRetrabalho = "select a.idprogramacao,b.idcd,b.idretrabalho from tbMPItens as a inner join tbRetrabalho as b on a.idprogramacao = b.idprogramacao where codigobarra = '" & vCodBarra & "'"
    rsLocalizaRetrabalho.Open SqlLocalizaRetrabalho, cnBanco, adOpenKeyset, adLockReadOnly
    If rsLocalizaRetrabalho.RecordCount > 0 Then
        Dim rsAlteraStatusRNC As New ADODB.Recordset
        Dim SqlAlteraStatusRNC As String
    
        Dim rsAlteraStatusCD As New ADODB.Recordset
        Dim SqlAlteraStatusCD As String
    
        SqlAlteraStatusRNC = "update tbRNC set status = 10 where idcd = '" & rsLocalizaRetrabalho.Fields(1) & "'"
        rsAlteraStatusRNC.Open SqlAlteraStatusRNC, cnBanco
        
        SqlAlteraStatusRNC = "update tbRNC set datareinsp = GETDATE() where idcd = '" & rsLocalizaRetrabalho.Fields(1) & "'"
        rsAlteraStatusRNC.Open SqlAlteraStatusRNC, cnBanco
        
        SqlAlteraStatusCD = "update tbComunicacaoDesvio set status = 10 where idcd = '" & rsLocalizaRetrabalho.Fields(1) & "'"
        rsAlteraStatusCD.Open SqlAlteraStatusCD, cnBanco
    End If
    rsLocalizaRetrabalho.Close
    Set rsLocalizaRetrabalho = Nothing
End Sub

Private Sub RestauraEmailEnvio(vModulo, vRecEmails As String)
    Dim rsEmailSystem As New ADODB.Recordset
    Dim sqlEmailSystem As String
    Dim X As Integer
    sqlEmailSystem = "Select * from tbEnvioEmail where modulo = '" & vModulo & "'"
    rsEmailSystem.Open sqlEmailSystem, cnBanco, adOpenKeyset, adLockReadOnly
    For X = 1 To rsEmailSystem.RecordCount
        If vRecEmails = "" Then
            vRecEmails = rsEmailSystem.Fields(1)
        Else
            vRecEmails = vRecEmails & ";" & rsEmailSystem.Fields(1)
        End If
        rsEmailSystem.MoveNext
    Next
    rsEmailSystem.Close
    Set rsEmailSystem = Nothing
End Sub

