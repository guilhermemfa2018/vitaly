VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AlteraClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTERA CADASTRO DO CLIENTE"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   Icon            =   "AlteraClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7335
      Index           =   0
      Left            =   1920
      TabIndex        =   139
      Top             =   360
      Width           =   10335
      Begin VB.CheckBox Funcionario 
         Caption         =   "Funcionário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox Fornecedor 
         Caption         =   "Fornecedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "AlteraClientes.frx":0E42
         Left            =   7800
         List            =   "AlteraClientes.frx":0E44
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox LimiteCredito 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   8520
         TabIndex        =   3
         Top             =   720
         Width           =   1665
      End
      Begin VB.TextBox Codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   450
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   200
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         FillColor       =   &H80000001&
         ForeColor       =   &H80000007&
         Height          =   2535
         Left            =   7920
         ScaleHeight     =   2505
         ScaleWidth      =   2265
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2295
         Begin VB.Image Image1 
            Height          =   2430
            Left            =   45
            Stretch         =   -1  'True
            Top             =   45
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdExcFoto 
         BackColor       =   &H80000003&
         Caption         =   "Excluir Foto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   144
         Top             =   4800
         Width           =   2295
      End
      Begin VB.CommandButton cmdFoto 
         BackColor       =   &H80000003&
         Caption         =   "Procurar Foto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   143
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox Situaçao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "AlteraClientes.frx":0E46
         Left            =   8520
         List            =   "AlteraClientes.frx":0E53
         TabIndex        =   2
         Text            =   "Ativo"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox RazaoSocial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   5865
      End
      Begin VB.TextBox Nome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   8
         Top             =   1680
         Width           =   5865
      End
      Begin VB.TextBox Rg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   6000
         TabIndex        =   11
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox Cliente 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox TipoPessoa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "AlteraClientes.frx":0E72
         Left            =   1800
         List            =   "AlteraClientes.frx":0E7C
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox InscEst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   6000
         TabIndex        =   13
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox PontoRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   6000
         TabIndex        =   26
         Top             =   6600
         Width           =   4185
      End
      Begin VB.TextBox Cidade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   6960
         TabIndex        =   23
         Top             =   6120
         Width           =   2745
      End
      Begin VB.TextBox Bairro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   22
         Top             =   6120
         Width           =   3585
      End
      Begin VB.TextBox Rua 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   20
         Top             =   5640
         Width           =   5865
      End
      Begin VB.TextBox Estado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   24
         Top             =   6600
         Width           =   585
      End
      Begin VB.TextBox Numero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   9360
         TabIndex        =   21
         Top             =   5640
         Width           =   825
      End
      Begin VB.CommandButton cmdRuas 
         BackColor       =   &H8000000A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7800
         TabIndex        =   142
         ToolTipText     =   "Cadastro de Ruas"
         Top             =   5640
         Width           =   375
      End
      Begin VB.CommandButton cmdBairros 
         BackColor       =   &H8000000A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         TabIndex        =   141
         ToolTipText     =   "Cadastro de Bairros"
         Top             =   6120
         Width           =   375
      End
      Begin VB.CommandButton cmdCidades 
         BackColor       =   &H8000000A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9840
         TabIndex        =   140
         ToolTipText     =   "Cadastro de Cidades"
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox LocalNasc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   19
         Top             =   5040
         Width           =   5865
      End
      Begin VB.TextBox EstadoCivil 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1800
         TabIndex        =   18
         Top             =   4560
         Width           =   2265
      End
      Begin MSMask.MaskEdBox CPF 
         Height          =   360
         Left            =   1800
         TabIndex        =   10
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   " "
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   0
         Left            =   600
         OleObjectBlob   =   "AlteraClientes.frx":0EA0
         TabIndex        =   147
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Telefone 
         Height          =   360
         Left            =   6000
         TabIndex        =   15
         Top             =   3600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fax 
         Height          =   360
         Left            =   6000
         TabIndex        =   17
         Top             =   4080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   " "
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Index           =   0
         Left            =   5040
         OleObjectBlob   =   "AlteraClientes.frx":0F06
         TabIndex        =   148
         Top             =   3600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Index           =   0
         Left            =   5160
         OleObjectBlob   =   "AlteraClientes.frx":0F6C
         TabIndex        =   149
         Top             =   4080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":0FD0
         TabIndex        =   150
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Index           =   0
         Left            =   4200
         OleObjectBlob   =   "AlteraClientes.frx":1042
         TabIndex        =   151
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox CNPJ 
         Height          =   360
         Left            =   1800
         TabIndex        =   12
         Top             =   3120
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
         Height          =   255
         Index           =   0
         Left            =   720
         OleObjectBlob   =   "AlteraClientes.frx":10B2
         TabIndex        =   152
         Top             =   3120
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Index           =   1
         Left            =   5040
         OleObjectBlob   =   "AlteraClientes.frx":1118
         TabIndex        =   153
         Top             =   2640
         Width           =   855
      End
      Begin MSMask.MaskEdBox Celular 
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   " "
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Index           =   1
         Left            =   960
         OleObjectBlob   =   "AlteraClientes.frx":117A
         TabIndex        =   154
         Top             =   4080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":11E6
         TabIndex        =   155
         Top             =   3600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Index           =   1
         Left            =   1080
         OleObjectBlob   =   "AlteraClientes.frx":1256
         TabIndex        =   156
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
         Height          =   255
         Index           =   1
         Left            =   720
         OleObjectBlob   =   "AlteraClientes.frx":12BC
         TabIndex        =   157
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Index           =   0
         Left            =   4680
         OleObjectBlob   =   "AlteraClientes.frx":1320
         TabIndex        =   158
         Top             =   3120
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Index           =   2
         Left            =   7320
         OleObjectBlob   =   "AlteraClientes.frx":1392
         TabIndex        =   159
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Index           =   2
         Left            =   6480
         OleObjectBlob   =   "AlteraClientes.frx":1400
         TabIndex        =   160
         Top             =   720
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Index           =   4
         Left            =   6480
         OleObjectBlob   =   "AlteraClientes.frx":1480
         TabIndex        =   161
         Top             =   1200
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DataCadastro 
         Height          =   360
         Left            =   5400
         TabIndex        =   1
         Top             =   240
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   0
         CheckBox        =   -1  'True
         Format          =   63242241
         CurrentDate     =   38656
      End
      Begin MSComCtl2.DTPicker DataNasc 
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   3600
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CheckBox        =   -1  'True
         Format          =   63242241
         CurrentDate     =   38656
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Index           =   0
         Left            =   1080
         OleObjectBlob   =   "AlteraClientes.frx":14EE
         TabIndex        =   162
         Top             =   5640
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   0
         Left            =   5880
         OleObjectBlob   =   "AlteraClientes.frx":1552
         TabIndex        =   163
         Top             =   6120
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Index           =   0
         Left            =   960
         OleObjectBlob   =   "AlteraClientes.frx":15BC
         TabIndex        =   164
         Top             =   6600
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Index           =   0
         Left            =   2520
         OleObjectBlob   =   "AlteraClientes.frx":1626
         TabIndex        =   165
         Top             =   6600
         Width           =   495
      End
      Begin MSMask.MaskEdBox Cep 
         Height          =   360
         Left            =   3120
         TabIndex        =   25
         Top             =   6600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###-###"
         PromptChar      =   " "
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   1
         Left            =   840
         OleObjectBlob   =   "AlteraClientes.frx":168A
         TabIndex        =   166
         Top             =   6120
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Index           =   1
         Left            =   8400
         OleObjectBlob   =   "AlteraClientes.frx":16F4
         TabIndex        =   167
         Top             =   5640
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Index           =   3
         Left            =   4680
         OleObjectBlob   =   "AlteraClientes.frx":175E
         TabIndex        =   168
         Top             =   6600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Index           =   3
         Left            =   360
         OleObjectBlob   =   "AlteraClientes.frx":17D0
         TabIndex        =   169
         Top             =   5040
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":1844
         TabIndex        =   170
         Top             =   4560
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Index           =   4
      Left            =   1920
      TabIndex        =   80
      Top             =   360
      Width           =   10335
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "AlteraClientes.frx":18BA
         Height          =   2655
         Left            =   360
         OleObjectBlob   =   "AlteraClientes.frx":18CE
         TabIndex        =   82
         Top             =   2160
         Width           =   9615
      End
      Begin VB.CommandButton cmdExcluirConta 
         BackColor       =   &H8000000A&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   79
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdInserirConta 
         BackColor       =   &H8000000A&
         Caption         =   "Inserir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   78
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Index           =   3
      Left            =   1920
      TabIndex        =   102
      Top             =   360
      Width           =   10335
      Begin VB.CommandButton cmdInserirRef 
         BackColor       =   &H8000000A&
         Caption         =   "Inserir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   62
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox EndereçoRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1440
         TabIndex        =   59
         Top             =   1320
         Width           =   5040
      End
      Begin VB.TextBox CidadeRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   6720
         TabIndex        =   58
         Top             =   840
         Width           =   3360
      End
      Begin VB.TextBox EstadoRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   7440
         TabIndex        =   60
         Top             =   1320
         Width           =   585
      End
      Begin VB.CommandButton cmdExcluirRef 
         BackColor       =   &H8000000A&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   116
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox NomeRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   5160
         TabIndex        =   55
         Top             =   360
         Width           =   4935
      End
      Begin VB.ComboBox Referencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "AlteraClientes.frx":381D
         Left            =   1440
         List            =   "AlteraClientes.frx":382D
         TabIndex        =   54
         Top             =   360
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Caption         =   "Local de Trabalho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2775
         Left            =   120
         TabIndex        =   103
         Top             =   4440
         Width           =   10215
         Begin VB.TextBox Contato 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   6960
            TabIndex        =   66
            Top             =   840
            Width           =   3120
         End
         Begin VB.TextBox Empresa 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1680
            TabIndex        =   63
            Top             =   360
            Width           =   5760
         End
         Begin VB.TextBox Ramal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1680
            TabIndex        =   65
            Top             =   840
            Width           =   4185
         End
         Begin VB.TextBox CidadeTrab 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   6960
            TabIndex        =   68
            Top             =   1320
            Width           =   3120
         End
         Begin VB.TextBox EnderecoTrab 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1680
            TabIndex        =   67
            Top             =   1320
            Width           =   4200
         End
         Begin VB.TextBox Salario 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   8400
            TabIndex        =   71
            Top             =   1800
            Width           =   1665
         End
         Begin VB.TextBox RegEmpregador 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   5640
            TabIndex        =   70
            Top             =   1800
            Width           =   1785
         End
         Begin VB.TextBox EmpregoAnterior 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1680
            TabIndex        =   72
            Top             =   2280
            Width           =   3840
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Index           =   0
            Left            =   6000
            OleObjectBlob   =   "AlteraClientes.frx":3859
            TabIndex        =   104
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Index           =   0
            Left            =   7440
            OleObjectBlob   =   "AlteraClientes.frx":38C5
            TabIndex        =   105
            Top             =   360
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Index           =   0
            Left            =   600
            OleObjectBlob   =   "AlteraClientes.frx":392B
            TabIndex        =   106
            Top             =   840
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Index           =   1
            Left            =   360
            OleObjectBlob   =   "AlteraClientes.frx":3995
            TabIndex        =   107
            Top             =   360
            Width           =   1215
         End
         Begin MSMask.MaskEdBox TelefoneComercial 
            Height          =   360
            Left            =   8280
            TabIndex        =   64
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   4210752
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)####-####"
            PromptChar      =   " "
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "AlteraClientes.frx":3A01
            TabIndex        =   108
            Top             =   1320
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   6000
            OleObjectBlob   =   "AlteraClientes.frx":3A6F
            TabIndex        =   109
            Top             =   1320
            Width           =   855
         End
         Begin MSComCtl2.DTPicker AdmitidoEm 
            Height          =   360
            Left            =   1680
            TabIndex        =   69
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   4210752
            CheckBox        =   -1  'True
            Format          =   63242241
            CurrentDate     =   38656
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Index           =   3
            Left            =   240
            OleObjectBlob   =   "AlteraClientes.frx":3AD9
            TabIndex        =   110
            Top             =   1800
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   7440
            OleObjectBlob   =   "AlteraClientes.frx":3B4D
            TabIndex        =   111
            Top             =   1800
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
            Height          =   255
            Left            =   3600
            OleObjectBlob   =   "AlteraClientes.frx":3BB9
            TabIndex        =   112
            Top             =   1800
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Index           =   4
            Left            =   120
            OleObjectBlob   =   "AlteraClientes.frx":3C35
            TabIndex        =   113
            Top             =   2280
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Index           =   1
            Left            =   7800
            OleObjectBlob   =   "AlteraClientes.frx":3CAB
            TabIndex        =   114
            Top             =   2280
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Index           =   2
            Left            =   5520
            OleObjectBlob   =   "AlteraClientes.frx":3D0F
            TabIndex        =   115
            Top             =   2280
            Width           =   495
         End
         Begin MSComCtl2.DTPicker DataIniEmpAnterior 
            Height          =   360
            Left            =   6120
            TabIndex        =   73
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   4210752
            CheckBox        =   -1  'True
            Format          =   63242241
            CurrentDate     =   38656
         End
         Begin MSComCtl2.DTPicker DatafimEmpAnterior 
            Height          =   360
            Left            =   8400
            TabIndex        =   74
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   4210752
            CheckBox        =   -1  'True
            Format          =   63242241
            CurrentDate     =   38656
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "AlteraClientes.frx":3D71
         TabIndex        =   117
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":3DE3
         TabIndex        =   118
         Top             =   1320
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "AlteraClientes.frx":3E51
         TabIndex        =   119
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "AlteraClientes.frx":3EBB
         TabIndex        =   120
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   8160
         OleObjectBlob   =   "AlteraClientes.frx":3F25
         TabIndex        =   121
         Top             =   1320
         Width           =   495
      End
      Begin MSMask.MaskEdBox CEPRef 
         Height          =   360
         Left            =   8760
         TabIndex        =   61
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###-###"
         PromptChar      =   " "
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "AlteraClientes.frx":3F89
         Height          =   2055
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":3F9D
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   2280
         Width           =   9855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "AlteraClientes.frx":5860
         TabIndex        =   123
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":58C6
         TabIndex        =   124
         Top             =   840
         Width           =   1095
      End
      Begin MSMask.MaskEdBox TelefoneRef 
         Height          =   360
         Left            =   1440
         TabIndex        =   56
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   " "
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "AlteraClientes.frx":5934
         TabIndex        =   125
         Top             =   840
         Width           =   855
      End
      Begin MSMask.MaskEdBox CelularRef 
         Height          =   360
         Left            =   4080
         TabIndex        =   57
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Index           =   2
      Left            =   1920
      TabIndex        =   126
      Top             =   360
      Width           =   10335
      Begin VB.CommandButton cmdVerDep 
         BackColor       =   &H8000000A&
         Caption         =   "Visualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   173
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdLimpaDep 
         BackColor       =   &H8000000A&
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   172
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Parentesco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1560
         TabIndex        =   51
         Top             =   2880
         Width           =   3945
      End
      Begin VB.CommandButton cmdInserirFilho 
         BackColor       =   &H8000000A&
         Caption         =   "Inserir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   52
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdExcluirFilho 
         BackColor       =   &H8000000A&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   127
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Filho 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1560
         TabIndex        =   41
         Top             =   480
         Width           =   5340
      End
      Begin VB.TextBox RgDep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1560
         TabIndex        =   43
         Top             =   960
         Width           =   2145
      End
      Begin VB.TextBox TrabalhoDep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   5670
         TabIndex        =   44
         Top             =   960
         Width           =   4425
      End
      Begin VB.TextBox CidadeTrabDep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1560
         TabIndex        =   46
         Top             =   1920
         Width           =   3825
      End
      Begin VB.TextBox EndTrabDep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1560
         TabIndex        =   45
         Top             =   1440
         Width           =   8505
      End
      Begin VB.TextBox RegEmpregadorDep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   7560
         TabIndex        =   47
         Top             =   1920
         Width           =   2505
      End
      Begin VB.TextBox SalarioDep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   7560
         TabIndex        =   50
         Top             =   2400
         Width           =   2505
      End
      Begin VB.CommandButton cmdAlterarFilho 
         BackColor       =   &H8000000A&
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   53
         Top             =   3480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "AlteraClientes.frx":59A0
         TabIndex        =   128
         Top             =   1440
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DataNascFilho 
         Height          =   360
         Left            =   8310
         TabIndex        =   42
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   4210752
         CheckBox        =   -1  'True
         Format          =   63242241
         CurrentDate     =   38656
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Index           =   1
         Left            =   6960
         OleObjectBlob   =   "AlteraClientes.frx":5A0E
         TabIndex        =   129
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Index           =   1
         Left            =   840
         OleObjectBlob   =   "AlteraClientes.frx":5A7E
         TabIndex        =   130
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "AlteraClientes.frx":5AE4
         TabIndex        =   131
         Top             =   960
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "AlteraClientes.frx":5B46
         TabIndex        =   132
         Top             =   960
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "AlteraClientes.frx":5BC0
         TabIndex        =   133
         Top             =   1920
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "AlteraClientes.frx":5C2A
         TabIndex        =   134
         Top             =   2400
         Width           =   735
      End
      Begin MSMask.MaskEdBox FoneDep 
         Height          =   360
         Left            =   4440
         TabIndex        =   49
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   4210752
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.DTPicker AdmitidoEmDep 
         Height          =   360
         Left            =   1560
         TabIndex        =   48
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   4210752
         CheckBox        =   -1  'True
         Format          =   63242241
         CurrentDate     =   38656
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "AlteraClientes.frx":5C90
         TabIndex        =   135
         Top             =   2400
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "AlteraClientes.frx":5D04
         TabIndex        =   136
         Top             =   2400
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "AlteraClientes.frx":5D70
         TabIndex        =   137
         Top             =   1920
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "AlteraClientes.frx":5DEC
         TabIndex        =   138
         Top             =   2880
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "AlteraClientes.frx":5E5E
         Height          =   3135
         Left            =   240
         OleObjectBlob   =   "AlteraClientes.frx":5E72
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   3960
         Width           =   9855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Index           =   1
      Left            =   1920
      TabIndex        =   83
      Top             =   360
      Width           =   10335
      Begin VB.Frame Frame2 
         Caption         =   "Outros Endereços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4935
         Left            =   120
         TabIndex        =   89
         Top             =   2280
         Width           =   10095
         Begin VB.TextBox Cidade1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   6480
            TabIndex        =   36
            Top             =   840
            Width           =   2985
         End
         Begin VB.TextBox Bairro1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1080
            TabIndex        =   35
            Top             =   840
            Width           =   3840
         End
         Begin VB.TextBox Rua1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1080
            TabIndex        =   33
            Top             =   360
            Width           =   6120
         End
         Begin VB.CommandButton cmdExcluirEndereço 
            BackColor       =   &H8000000A&
            Caption         =   "Excluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   93
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdInserirEndereço 
            BackColor       =   &H8000000A&
            Caption         =   "Inserir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   40
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox TipoEndereço 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   5640
            TabIndex        =   39
            Top             =   1320
            Width           =   4320
         End
         Begin VB.TextBox Estado1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   1080
            TabIndex        =   37
            Top             =   1320
            Width           =   585
         End
         Begin VB.TextBox Numero1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   9000
            TabIndex        =   34
            Top             =   360
            Width           =   945
         End
         Begin VB.CommandButton cmdCidades1 
            BackColor       =   &H8000000A&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9600
            TabIndex        =   92
            ToolTipText     =   "Cadastro de Cidades"
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdBairros1 
            BackColor       =   &H8000000A&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5040
            TabIndex        =   91
            ToolTipText     =   "Cadastro de Bairros"
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdRuas1 
            BackColor       =   &H8000000A&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7320
            TabIndex        =   90
            ToolTipText     =   "Cadastro de Ruas"
            Top             =   360
            Width           =   375
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "AlteraClientes.frx":7C6D
            Height          =   2535
            Left            =   120
            OleObjectBlob   =   "AlteraClientes.frx":7C81
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   2280
            Width           =   9855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   4920
            OleObjectBlob   =   "AlteraClientes.frx":9384
            TabIndex        =   95
            Top             =   1320
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "AlteraClientes.frx":93EA
            TabIndex        =   96
            Top             =   360
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   5520
            OleObjectBlob   =   "AlteraClientes.frx":944E
            TabIndex        =   97
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "AlteraClientes.frx":94B8
            TabIndex        =   98
            Top             =   1320
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   1920
            OleObjectBlob   =   "AlteraClientes.frx":9522
            TabIndex        =   99
            Top             =   1320
            Width           =   495
         End
         Begin MSMask.MaskEdBox Cep1 
            Height          =   360
            Left            =   2520
            TabIndex        =   38
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   4210752
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###-###"
            PromptChar      =   " "
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   8040
            OleObjectBlob   =   "AlteraClientes.frx":9586
            TabIndex        =   100
            Top             =   360
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "AlteraClientes.frx":95F0
            TabIndex        =   101
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Endereço Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1935
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   10095
         Begin VB.TextBox EndAnterior 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   3960
            TabIndex        =   32
            Top             =   1320
            Width           =   5880
         End
         Begin VB.TextBox ValorAluguel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   3960
            TabIndex        =   30
            Top             =   840
            Width           =   2145
         End
         Begin VB.TextBox TempoEnd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   3960
            TabIndex        =   27
            Top             =   360
            Width           =   2520
         End
         Begin VB.OptionButton OptCasaPropria 
            Caption         =   "Casa Própria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   29
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptAluguel 
            Caption         =   "Aluguel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Pensao 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   7320
            TabIndex        =   31
            Top             =   840
            Width           =   2505
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Index           =   1
            Left            =   1200
            OleObjectBlob   =   "AlteraClientes.frx":965A
            TabIndex        =   85
            Top             =   840
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Index           =   3
            Left            =   240
            OleObjectBlob   =   "AlteraClientes.frx":96E2
            TabIndex        =   86
            Top             =   360
            Width           =   3615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "AlteraClientes.frx":977C
            TabIndex        =   87
            Top             =   1320
            Width           =   3615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
            Height          =   255
            Left            =   6360
            OleObjectBlob   =   "AlteraClientes.frx":981A
            TabIndex        =   88
            Top             =   840
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdLimpar 
      BackColor       =   &H8000000A&
      Caption         =   "Limpar (F11)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   76
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdGravar 
      BackColor       =   &H80000003&
      Caption         =   "Gravar (F9)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   75
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H8000000A&
      Caption         =   "Sair (Esc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   77
      Top             =   1680
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7815
      Left            =   1800
      TabIndex        =   81
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13785
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cliente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dependentes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Referências"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contas"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "AlteraClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Filho_GotFocus()
Frame1(2).ZOrder
TabStrip1.Tabs(3).Selected = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then 'F9
    'cmdGravar_Click
ElseIf KeyCode = 122 Then 'F11
    'cmdLimpar_Click
End If
End Sub

Private Sub Form_Load()
AplicarSkin Me, Principal.Skin1
NewColorDBGrid Me

Call BordasControle(Me, DBGrid1, False)
Call BordasControle(Me, DBGrid2, False)
Call BordasControle(Me, DBGrid3, False)

Frame1(0).ZOrder
AjustaContainerClientes

End Sub

Private Sub AjustaContainerClientes()
Dim i As Integer
With TabStrip1
For i = 1 To .Tabs.Count
Frame1(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Next
End With
End Sub

Private Sub Referencia_GotFocus()
Frame1(3).ZOrder
TabStrip1.Tabs(4).Selected = True
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer

i = TabStrip1.SelectedItem.Index

Frame1(i - 1).ZOrder
End Sub

Private Sub TempoEnd_GotFocus()
Frame1(1).ZOrder
TabStrip1.Tabs(2).Selected = True
End Sub

Private Sub TipoPessoa_GotFocus()
Frame1(0).ZOrder
TabStrip1.Tabs(1).Selected = True
End Sub
