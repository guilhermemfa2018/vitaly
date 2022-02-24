VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmRelExpAvulso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de relatórios de expedição de terceiros"
   ClientHeight    =   8460
   ClientLeft      =   45
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
   Icon            =   "frmRelExpAvulso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
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
      Picture         =   "frmRelExpAvulso.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7560
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
      Picture         =   "frmRelExpAvulso.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "Salvar Relatório"
      ToolTipText     =   "Salvar Relatório"
      Top             =   7560
      Width           =   615
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
      TabIndex        =   64
      Top             =   120
      Width           =   2775
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":265E
         TabIndex        =   2
         Top             =   240
         Width           =   2535
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
      TabIndex        =   63
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
         Format          =   288489473
         CurrentDate     =   40449
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do destinatário"
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
      TabIndex        =   57
      Top             =   960
      Width           =   6975
      Begin VB.CommandButton cmdExpAvulso 
         Caption         =   "..."
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   71
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Tag             =   "Projeto"
         ToolTipText     =   "Projeto"
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton cmdExpAvulso 
         Caption         =   "..."
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
         Index           =   1
         Left            =   6480
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmRelExpAvulso.frx":26D2
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExpAvulso 
         Caption         =   "..."
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
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   6735
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6000
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "FCE nº"
         ToolTipText     =   "FCE nº"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   6255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":273A
         TabIndex        =   59
         Top             =   1440
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":27A8
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   6000
         OleObjectBlob   =   "frmRelExpAvulso.frx":280E
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":287C
         TabIndex        =   62
         Top             =   840
         Width           =   1215
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
      TabIndex        =   56
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1935
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
      TabIndex        =   40
      Top             =   3240
      Width           =   6975
      Begin VB.CommandButton cmdExpAvulso 
         Caption         =   "..."
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
         Index           =   2
         Left            =   6480
         TabIndex        =   11
         Top             =   480
         Width           =   375
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
         Height          =   975
         Left            =   120
         TabIndex        =   50
         Top             =   2760
         Width           =   6735
         Begin VB.ComboBox cboCadastro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            ItemData        =   "frmRelExpAvulso.frx":28E4
            Left            =   1440
            List            =   "frmRelExpAvulso.frx":2939
            TabIndex        =   20
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cboCadastro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            ItemData        =   "frmRelExpAvulso.frx":29A9
            Left            =   3720
            List            =   "frmRelExpAvulso.frx":29FE
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   14
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   15
            Left            =   2280
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   16
            Left            =   4560
            TabIndex        =   23
            Top             =   480
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   4560
            OleObjectBlob   =   "frmRelExpAvulso.frx":2A6E
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmRelExpAvulso.frx":2ADA
            TabIndex        =   52
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   2280
            OleObjectBlob   =   "frmRelExpAvulso.frx":2B3E
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   1440
            OleObjectBlob   =   "frmRelExpAvulso.frx":2BB8
            TabIndex        =   54
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRelExpAvulso.frx":2C1C
            TabIndex        =   55
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtcadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Tag             =   "Código da transportadora"
         ToolTipText     =   "Código da transportadora"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   3120
         TabIndex        =   17
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   5880
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cboCadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmRelExpAvulso.frx":2C94
         Left            =   6120
         List            =   "frmRelExpAvulso.frx":2CE9
         TabIndex        =   18
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   3360
         TabIndex        =   13
         Top             =   1080
         Width           =   3495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "frmRelExpAvulso.frx":2D59
         TabIndex        =   41
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   6120
         OleObjectBlob   =   "frmRelExpAvulso.frx":2DB9
         TabIndex        =   42
         Top             =   2040
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRelExpAvulso.frx":2E17
         TabIndex        =   43
         Top             =   2040
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":2E7D
         TabIndex        =   44
         Top             =   2040
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":2EE3
         TabIndex        =   45
         Top             =   1440
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "frmRelExpAvulso.frx":2F4D
         TabIndex        =   46
         Top             =   840
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":2FC3
         TabIndex        =   47
         Top             =   840
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmRelExpAvulso.frx":3025
         TabIndex        =   48
         Top             =   240
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":3087
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
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
      TabIndex        =   36
      Top             =   120
      Width           =   14175
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmRelExpAvulso.frx":30ED
         TabIndex        =   38
         Top             =   7800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmRelExpAvulso.frx":314D
         TabIndex        =   39
         Top             =   7800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   12000
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
         Height          =   255
         Left            =   12000
         OleObjectBlob   =   "frmRelExpAvulso.frx":31C7
         TabIndex        =   98
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   5760
         TabIndex        =   27
         Top             =   1680
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "frmRelExpAvulso.frx":3235
         TabIndex        =   97
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtcadastro 
         Height          =   285
         Index           =   19
         Left            =   6360
         TabIndex        =   96
         Tag             =   "Código do Material"
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdExpAvulso 
         Caption         =   "..."
         Height          =   255
         Index           =   8
         Left            =   11400
         TabIndex        =   95
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Digitar"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   94
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Buscar"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   93
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados do destinatário avulso "
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
         Left            =   240
         TabIndex        =   72
         Top             =   5400
         Visible         =   0   'False
         Width           =   11415
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   9
            Left            =   5640
            TabIndex        =   92
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   8
            Left            =   10800
            TabIndex        =   91
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   7
            Left            =   8640
            TabIndex        =   90
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   6
            Left            =   6480
            TabIndex        =   89
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   5
            Left            =   5640
            TabIndex        =   88
            Top             =   1080
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "frmRelExpAvulso.frx":329D
            TabIndex        =   87
            Top             =   1440
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
            Height          =   255
            Left            =   10800
            OleObjectBlob   =   "frmRelExpAvulso.frx":330D
            TabIndex        =   86
            Top             =   840
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   8640
            OleObjectBlob   =   "frmRelExpAvulso.frx":3371
            TabIndex        =   85
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   6720
            OleObjectBlob   =   "frmRelExpAvulso.frx":33DD
            TabIndex        =   84
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "frmRelExpAvulso.frx":3449
            TabIndex        =   83
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   4
            Left            =   2880
            TabIndex        =   82
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   3
            Left            =   120
            TabIndex        =   81
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   80
            Top             =   1080
            Width           =   5415
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   1
            Left            =   1320
            TabIndex        =   79
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   78
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRelExpAvulso.frx":34AF
            TabIndex        =   77
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmRelExpAvulso.frx":350D
            TabIndex        =   76
            Top             =   1440
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRelExpAvulso.frx":356B
            TabIndex        =   75
            Top             =   1440
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRelExpAvulso.frx":35CD
            TabIndex        =   74
            Top             =   840
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmRelExpAvulso.frx":3637
            TabIndex        =   73
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdExpAvulso 
         Height          =   615
         Index           =   6
         Left            =   1920
         Picture         =   "frmRelExpAvulso.frx":3699
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "Excluir"
         ToolTipText     =   "Excluir"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdExpAvulso 
         Height          =   615
         Index           =   5
         Left            =   1320
         Picture         =   "frmRelExpAvulso.frx":4363
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "Editar"
         ToolTipText     =   "Editar"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdExpAvulso 
         Height          =   615
         Index           =   4
         Left            =   720
         Picture         =   "frmRelExpAvulso.frx":502D
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "Novo"
         ToolTipText     =   "Novo"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdExpAvulso 
         Height          =   615
         Index           =   3
         Left            =   120
         Picture         =   "frmRelExpAvulso.frx":5CF7
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "Incluir"
         ToolTipText     =   "Incluir"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmRelExpAvulso.frx":69C1
         Left            =   13200
         List            =   "frmRelExpAvulso.frx":69D7
         TabIndex        =   26
         Text            =   "PÇ"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   12000
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   11175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
         Height          =   255
         Left            =   13200
         OleObjectBlob   =   "frmRelExpAvulso.frx":69F3
         TabIndex        =   69
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
         Height          =   255
         Left            =   12000
         OleObjectBlob   =   "frmRelExpAvulso.frx":6A5B
         TabIndex        =   68
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRelExpAvulso.frx":6AC9
         TabIndex        =   67
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         Caption         =   "Item"
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
         Left            =   2760
         TabIndex        =   66
         Top             =   1440
         Width           =   1215
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   70
            Text            =   "-"
            Top             =   240
            Width           =   975
         End
      End
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
         TabIndex        =   37
         Top             =   7800
         Width           =   1000
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   9763
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
      Begin IpToolTips.cIpToolTips cIpToolTips1 
         Left            =   11280
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         BackColor       =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "frmRelExpAvulso.frx":6B35
         TabIndex        =   99
         Top             =   7800
         Visible         =   0   'False
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "frmRelExpAvulso.frx":6B95
         TabIndex        =   100
         Top             =   7800
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRelExpAvulso"
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
Private vPonte2 As TextBox
Private vPonte3 As TextBox
Private rsFCE As New ADODB.Recordset
Private sqlFCE As String
Private rsProjeto As New ADODB.Recordset
Private SqlProjeto As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 4
        'If ContaChecado > LimiTador Then
        '    mobjMsg.Abrir "Limite máximo de itens selecionados foi ultrapassado." & vbCrLf & "Limite Máximo: " & LimiTador, Ok, critico, "Atenção"
        'Else
            mobjMsg.Abrir "Deseja gravar o relatório?", YesNo, pergunta, "Zeus"
            If Tp = 1 Then
                If GravarDados = True Then
                    Unload Me
                End If
            End If
        'End If
    Case 6
        mobjMsg.Abrir "Deseja sair da tela de emissão de relatórios?", YesNo, pergunta, "Zeus"
        If Tp = 1 Then
            Unload Me
        End If
    End Select
End Sub

Private Sub cmdExpAvulso_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridFCE
        CarregaFCE
    Case 1
        If txtcadastro(0).Text <> "" Then
            ChamaGridProjeto
            CarregaProjeto
        End If
    Case 2
        ChamaGridTrans
        CarregaTipoTrans
        txtcadastro(14).SetFocus
    Case 3
        If Text2.Text = "" Then
            mobjMsg.Abrir "Favor informar a quantidade", Ok, critico, "Atenção"
            Exit Sub
        End If
        vPonte1.Text = Combo1.Text
        vPonte2.Text = " "
        vPonte3.Text = "12"
        If txtcadastro(19).Text <> "" Then Text1.Text = txtcadastro(19) & " - " & Text1.Text
        IncluirLV ListView1, Text3, Text1, Text4, vPonte2, vPonte2, Text5, vPonte2, Text2, vPonte3, vPonte1, vPonte2, vPonte2, vPonte2, vPonte2, vPonte2
        LimpaControles Text1, Text2, txtcadastro(19), Text4, Text5, Text1, Text1, Text1, Text1, Text1
        Text3 = Format(GeraCodigoLV(ListView1), "00")
        SomaTotais
        If Text1.Enabled = True Then Text1.SetFocus
    Case 4
        LimpaControles Text1, Text2, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1
        Text3 = Format(GeraCodigoLV(ListView1), "00")
    Case 5
        AlteraLV ListView1, Text3, Text1, Text4, vPonte3, vPonte3, Text5, vPonte3, Text2, vPonte3, vPonte1, vPonte3, vPonte3, vPonte3, vPonte3, vPonte3
        Combo1.Text = vPonte1.Text
    Case 6
        ExcluirItemLV ListView1
    Case 7
        ChamaGridCliFor
    Case 8
        ChamaGridProduto
        CarregaDados (19)
    End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    'AlwaysOnTop frmRelExpAvulso, True ' Mantem o formulário sempre em primeiro plano
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    Set vPonte3 = Me.Controls.Add("VB.TextBox", "vPonte3")
    Me.Top = 0
    Me.Left = (Principal.Width / 2) - (Me.Width / 2)
    
    ContaChecado = 0
    LimiTador = 1000 '49
    Legenda = "Aguarde"
    SelecionaLinha
    CompoeControles
    listview_cabecalho 'Chama a Sub que monta o cabeçalho das colunas do Listview
'
'    CompoeListview2 'Listview de Expedição

    txtLvw = ""
    txtLvw.Visible = False
    txtLvw.Tag = False 'is ListView2 dirty, not used in this example
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico

End Sub

Private Sub ListView1_DblClick()
    AlteraLV ListView1, Text3, Text1, vPonte3, vPonte3, vPonte3, vPonte3, vPonte3, Text2, vPonte3, vPonte1, vPonte3, vPonte3, vPonte3, vPonte3, vPonte3
    Combo1.Text = vPonte1.Text
End Sub

Private Sub optSelect_Click(Index As Integer)
    If optSelect(0).Value = True Then
        Text1.Enabled = False
        cmdExpAvulso(8).Enabled = True
    Else
        Text1.Enabled = True
        cmdExpAvulso(8).Enabled = False
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    'aceitar somente números e "Back Space", "Enter", "virgula"
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCadastro_GotFocus(Index As Integer)
    mudaCorText txtcadastro(Index)
End Sub

Private Sub Form_Resize()
    DimensionaFormExp frmRelExpAvulso
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
    ListView1.ColumnHeaders.Add , , "UN", ListView1.Width / 22
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
    txtcadastro(4) = Format(GeraCodigo, "000000000") & "" 'Identificador do relatório
    Text3 = Format(GeraCodigoLV(ListView1), "00")
    DTPicker1 = Date 'Data de emissão do relatório
    SkinLabel7.Caption = vSituacao
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
    SomaTotais = True
    Dim Y As Integer, SomaQtd As Double
    Y = ListView1.ListItems.Count
    SomaQtd = 0
    SomaPeso = 0
    For W = 1 To Y
        ListView1.ListItems(W).Selected = True
        SomaQtd = SomaQtd + ListView1.SelectedItem.ListSubItems.Item(7)
        SomaPeso = SomaPeso + ListView1.SelectedItem.ListSubItems.Item(5)
        ListView1.SelectedItem.ListSubItems.Item(12) = Format(0, "#,##0.00;(#,##0.00)")
    Next
    SkinLabel29 = Format(SomaQtd, "#,##0.00;(#,##0.00)")
    SkinLabel38 = Format(SomaPeso, "#,##0.00;(#,##0.00)")
    Exit Function
TrataErro:
    SomaTotais = False
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
            txtcadastro(6).Text = Format(rsLocal.Fields(0), "000")
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub CarregaTipoTrans()
On Error GoTo Err
    Dim X As Integer
    Dim rsTipoTrans As New ADODB.Recordset
    SqlM = "select a.CODTRA,a.NOME,a.CGC,a.INSCRESTADUAL,a.RUA+','+a.NUMERO as endereco,a.CEP,a.BAIRRO,a.CIDADE,a.CODETD from " & vBancoTotvs & ".dbo.ttra as a order by a.CODETD"
    rsTipoTrans.Open SqlM, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsTipoTrans.EOF Then rsTipoTrans.MoveFirst
    rsTipoTrans.Find "CODTRA=" & "'" & Format(txtcadastro(6), "000") & "'"
    If rsTipoTrans.EOF Then
        txtcadastro(6).Text = Format(txtcadastro(6), "000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Transportadora não cadastrada", Ok, critico, "Atenção"
            txtcadastro(7) = ""
        End If
    Else
        txtcadastro(6).Text = Format(rsTipoTrans.Fields(0), "000") & "" 'codigo
        txtcadastro(7).Text = rsTipoTrans.Fields(1) 'nome
        If Not IsNull(rsTipoTrans.Fields(2)) Then txtcadastro(8).Text = rsTipoTrans.Fields(2) 'cnpj
        If Not IsNull(rsTipoTrans.Fields(3)) Then txtcadastro(9).Text = rsTipoTrans.Fields(3) 'ie
        If Not IsNull(rsTipoTrans.Fields(4)) Then txtcadastro(10).Text = rsTipoTrans.Fields(4) 'endereco (rua+numero)
        If Not IsNull(rsTipoTrans.Fields(5)) Then txtcadastro(11).Text = rsTipoTrans.Fields(5) 'cep
        If Not IsNull(rsTipoTrans.Fields(6)) Then txtcadastro(12).Text = rsTipoTrans.Fields(6) 'bairro
        If Not IsNull(rsTipoTrans.Fields(7)) Then txtcadastro(13).Text = rsTipoTrans.Fields(7) 'cidade
        If Not IsNull(rsTipoTrans.Fields(8)) Then cboCadastro(0).Text = rsTipoTrans.Fields(8) 'UF
        For X = 7 To 13
            txtcadastro(X).Enabled = False
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

Private Sub ChamaGridCliFor()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "select A.CODCFO,a.NOME,B.DESCRICAO +': '+ A.RUA +', '+A.NUMERO AS ENDERECO,A.CEP,A.BAIRRO,A.CIDADE,A.CODETD AS UF,A.TELEFONE,A.CGCCFO,A.INSCRESTADUAL from " & vBancoTotvs & ".dbo.FCFO AS A LEFT JOIN " & vBancoTotvs & ".dbo.DTIPORUA AS B ON A.TIPORUA = B.CODIGO WHERE A.ATIVO = 1 order by a.nome"
    procnom = "CODCFO"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Destinatários"
    Pesquisa = frmRelExpAvulso.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        'Pesquisa = Mid$(Pesquisa, 7, 100)
        rsLocal.Find "CODCFO=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(0).Text = "0"
            txtcadastro(1).Text = "0"
            txtcadastro(2).Text = "-"
            txtcadastro(3).Text = rsLocal.Fields(1)

            txtDestinatario(0).Text = rsLocal.Fields(0) 'Identificador do destinatário
            txtDestinatario(1).Text = rsLocal.Fields(1) 'Nome do destinatário
            If Not IsNull(rsLocal.Fields(2)) Then txtDestinatario(2).Text = rsLocal.Fields(2) Else txtDestinatario(2).Text = "-" 'Endereço do destinatário
            If Not IsNull(rsLocal.Fields(8)) Then txtDestinatario(3).Text = rsLocal.Fields(8) 'CNPJ do destinatário
            txtDestinatario(4).Text = rsLocal.Fields(9) 'IE do destinatário

            If Not IsNull(rsLocal.Fields(3)) Then txtDestinatario(5).Text = rsLocal.Fields(3) 'CEP do destinatário
            txtDestinatario(6).Text = rsLocal.Fields(4) 'Bairro do destinatário
            txtDestinatario(7).Text = rsLocal.Fields(5) 'Cidade do destinatário
            txtDestinatario(8).Text = rsLocal.Fields(6) 'UF do destinatário
            If Not IsNull(rsLocal.Fields(7)) Then txtDestinatario(9).Text = rsLocal.Fields(7) 'Telefone do destinatário
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Function GravarDados()
On Error GoTo Err
    If ValidaCampo = False Then Exit Function
    GravarDados = True
    Dim Y As Integer, X As Integer
    
    Dim rsRelatorio As New ADODB.Recordset
    Dim sqlRelatorio As String
    Dim rsItensRelatorio As New ADODB.Recordset
    Dim sqlItensRelatorio As String
    
    
    'If SomaTotais = False Then Exit Sub
    If Val(SkinLabel29) = 0 Then
        mobjMsg.Abrir "Os campos referentes a quantidade estão vazios", Ok, critico, "Atenção"
        GravarDados = False
        Exit Function
    End If
    
10  cnBanco.BeginTrans

    sqlRelatorio = "select * from tbRelInspExp"
    rsRelatorio.Open sqlRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    rsRelatorio.AddNew
    
    txtcadastro(4) = Format(GeraCodigo, "000000000") & "" 'Identificador do relatório
    
    rsRelatorio.Fields(0) = Val(txtcadastro(4)) 'Codigo do Relatorio
    rsRelatorio.Fields(1) = Val(txtcadastro(0)) 'FCE
    rsRelatorio.Fields(2) = Val(txtcadastro(1)) 'Codigo do projeto
    rsRelatorio.Fields(3) = Format(DTPicker1, "dd/mm/yyyy") 'Data do relatorio
    rsRelatorio.Fields(4) = txtcadastro(5) 'Observação
    rsRelatorio.Fields(5) = 0 'Status de impressão
    'rsRelatorio.Fields(6) = cboCadastro(4) 'Norma de Liberação
    rsRelatorio.Fields(7) = 11 'Tipo relatorio (11) Expedição
    rsRelatorio.Fields(8) = Format(SkinLabel38, "#,##0.00;(#,##0.00)") 'Peso de balança
'    rsRelatorio.Fields(8) = Format(0, "#,##0.00;(#,##0.00)") 'Peso de balança
    rsRelatorio.Fields(9) = NomUsu
    
    If txtcadastro(0).Text = 0 Then
        rsRelatorio.Fields(10) = txtDestinatario(0) 'Identificador do Destinatário
        rsRelatorio.Fields(11) = txtDestinatario(1) 'nome do Destinatário
        rsRelatorio.Fields(12) = txtDestinatario(2) 'Endereço do Destinatário
        rsRelatorio.Fields(13) = txtDestinatario(3) 'CNPJ do Destinatário
        rsRelatorio.Fields(14) = txtDestinatario(4) 'IE do Destinatário
        rsRelatorio.Fields(15) = txtDestinatario(5) 'CEP do Destinatário
        rsRelatorio.Fields(16) = txtDestinatario(6) 'Bairro do Destinatário
        rsRelatorio.Fields(17) = txtDestinatario(7) 'Cidade do Destinatário
        rsRelatorio.Fields(18) = txtDestinatario(8) 'UF do Destinatário
        rsRelatorio.Fields(19) = txtDestinatario(9) 'Telefone do Destinatário
    End If
    rsRelatorio.Update
    rsRelatorio.Close
    Set rsRelatorio = Nothing

    'Gravar dados referente aos Itens do Relatório
    sqlItensRelatorio = "select * from tbRelInspExpitens"
    rsItensRelatorio.Open sqlItensRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        'If ListView1.ListItems.Item(X).Checked = True Then
            rsItensRelatorio.AddNew
            rsItensRelatorio.Fields(0) = Val(txtcadastro(4)) 'Codigo do relatorio
            rsItensRelatorio.Fields(1) = Val(txtcadastro(0).Text) 'Nº FCE
            rsItensRelatorio.Fields(2) = Val(txtcadastro(1)) 'Código do Projeto
            rsItensRelatorio.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2) 'Desenho
            rsItensRelatorio.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(3) 'Revisão do Desenho
            rsItensRelatorio.Fields(5) = ListView1.ListItems.Item(X) 'Posição
            rsItensRelatorio.Fields(6) = ListView1.SelectedItem.ListSubItems.Item(1) 'Descrição da posição
            rsItensRelatorio.Fields(7) = ListView1.SelectedItem.ListSubItems.Item(8) 'Status (Codfase)
            rsItensRelatorio.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(7) 'Quantidade liberada
            
            If SkinLabel7 = "EXPEDIÇÃO TERC." Then
                rsItensRelatorio.Fields(9) = ListView1.SelectedItem.ListSubItems.Item(5) 'Peso liberado
            Else
                rsItensRelatorio.Fields(9) = ListView1.SelectedItem.ListSubItems.Item(12) 'Peso liberado
            End If
            
            rsItensRelatorio.Fields(10) = Val(ListView1.SelectedItem.ListSubItems.Item(10)) 'Código da LM - Lista de Material
            rsItensRelatorio.Fields(11) = Val(ListView1.SelectedItem.ListSubItems.Item(11)) 'Código da sequencia da LM
            rsItensRelatorio.Fields(13) = "-" 'Inspeções realizadas
            rsItensRelatorio.Fields(14) = ListView1.SelectedItem.ListSubItems.Item(9) 'Unidade de medida
            rsItensRelatorio.Update
        'End If
    Next
    rsItensRelatorio.Close
    Set rsItemRelatorio = Nothing
    
    cnBanco.CommitTrans
    
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    If txtcadastro(0).Text = "0" Then vQualquerDado(20, 1) = "-" Else vQualquerDado(20, 1) = txtcadastro(0).Text 'grava o numero da FCE
    
    vQualquerDado(1, 1) = txtcadastro(4).Text
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = txtcadastro(6).Text
    vQualquerDado(2, 2) = "I"
    vQualquerDado(3, 1) = txtcadastro(7).Text
    vQualquerDado(3, 2) = "S"
    vQualquerDado(4, 1) = txtcadastro(8).Text
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtcadastro(9).Text
    vQualquerDado(5, 2) = "S"
    
    vQualquerDado(6, 1) = txtcadastro(10).Text
    vQualquerDado(6, 2) = "S"
    vQualquerDado(7, 1) = txtcadastro(11).Text
    vQualquerDado(7, 2) = "S"
    vQualquerDado(8, 1) = txtcadastro(12).Text
    vQualquerDado(8, 2) = "S"
    vQualquerDado(9, 1) = txtcadastro(13).Text
    vQualquerDado(9, 2) = "S"
    vQualquerDado(10, 1) = cboCadastro(0).Text
    vQualquerDado(10, 2) = "S"
    
    vQualquerDado(11, 1) = txtcadastro(14).Text
    vQualquerDado(11, 2) = "S"
    vQualquerDado(12, 1) = cboCadastro(1).Text
    vQualquerDado(12, 2) = "S"
    vQualquerDado(13, 1) = txtcadastro(15).Text
    vQualquerDado(13, 2) = "S"
    vQualquerDado(14, 1) = cboCadastro(2).Text
    vQualquerDado(14, 2) = "S"
    vQualquerDado(15, 1) = txtcadastro(16).Text
    vQualquerDado(15, 2) = "S"
    GravaDados "tbRelExpTransp", "codrel", "I", txtcadastro(4), 15, "", "", txtcadastro(4)
   
    mobjMsg.Abrir "Dados gravados com sucesso", Ok, informacao, "Atenção"
    
    mobjMsg.Abrir "Dados gravados com sucesso.Deseja imprimir de relatório de Expedição?", YesNo, pergunta, "Zeus"
    If Tp = 1 Then
        vCodRel = Val(txtcadastro(4))
        FCRExpedicao.Show 1
        sqlRelatorio = "update tbRelInspExp set statusimp=1 where codrel ='" & vCodRel & "'"
        rsRelatorio.Open sqlRelatorio, cnBanco, adOpenKeyset, adLockOptimistic
    End If
    
    Unload Me
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
        cnBanco.RollbackTrans
        GravarDados = False
        Exit Function
    End If
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If txtcadastro(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(0).Tag, Ok, critico, "Atenção"
        Me.cmdExpAvulso(0).SetFocus
        Exit Function
    End If
    If txtcadastro(2).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(2).Tag, Ok, critico, "Atenção"
        Me.cmdExpAvulso(1).SetFocus
        Exit Function
    End If
    If txtcadastro(6).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(6).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(6).SetFocus
        Exit Function
    End If
    If ListView1.ListItems.Count = 0 Then
        mobjMsg.Abrir "É necessário registrar ao menos 01 item para salvar o relatório", Ok, critico, "Atenção"
        Me.Text1.SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

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

Private Sub SelecionaLinha()
    Dim Y As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For W = 1 To Y
        If MeuLV.ListView1.ListItems.Item(W).Selected = True Then
            Exit For
        End If
    Next
    MeuLV.ListView1.ListItems(W).Selected = True
End Sub

Private Sub CarregaFCE()
On Error GoTo Err
    Dim X As Integer
    sqlFCE = "Select a.fce,a.codprojeto,b.codclifor,c.nome from tbprojetos as a inner join tbfo as b on a.fce = b.fce left join tbclifor as c on b.codclifor = c.codclifor where a.fce = '" & Val(txtcadastro(0)) & "' order by fce"
    rsFCE.Open sqlFCE, cnBanco, adOpenKeyset, adLockOptimistic
    If rsFCE.EOF Then
        txtcadastro(0).Text = txtcadastro(0)
        mobjMsg.Abrir "FCE não cadastrada", Ok, critico, "Atenção"
    Else
        txtcadastro(0).Text = rsFCE.Fields(0)
        txtcadastro(3).Text = rsFCE.Fields(3)
        txtcadastro(1).Text = ""
        txtcadastro(2).Text = ""
        txtDestinatario(0).Text = "" 'Identificador do destinatário
        txtDestinatario(1).Text = "" 'Nome do destinatário
        txtDestinatario(2).Text = "" 'Endereço do destinatário
        txtDestinatario(3).Text = "" 'CNPJ do destinatário
        txtDestinatario(4).Text = "" 'IE do destinatário
    End If
    rsFCE.Close
    Set rsFCE = Nothing
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

Private Sub ChamaGridFCE()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select B.fce,C.nome from tbFo AS B INNER JOIN tbclifor AS C ON B.codclifor = C.codclifor  order by b.fce"
    procnom = "FCE"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de FCE"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "fce=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(0).Text = rsLocal.Fields(0)
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub CarregaProjeto()
On Error GoTo Err
    Dim X As Integer
    SqlProjeto = "Select * from tbprojetos where fce = '" & txtcadastro(0) & "' order by fce"
    rsProjeto.Open SqlProjeto, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsProjeto.EOF Then rsProjeto.MoveFirst
    rsProjeto.Find "projeto=" & "'" & Me.txtcadastro(2) & "'"
    If rsProjeto.EOF Then
        txtcadastro(2).Text = txtcadastro(2)
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Projeto não cadastrado", Ok, critico, "Atenção"
        End If
    Else
        txtcadastro(2).Text = rsProjeto.Fields(2)
        txtcadastro(1).Text = rsProjeto.Fields(0)
        'txtDesenho(1).Text = Format(rsProjeto.Fields(0), "000000")
    End If
    rsProjeto.Close
    Set rsProjeto = Nothing
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

Private Sub ChamaGridProjeto()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbprojetos where fce = '" & txtcadastro(0) & "' order by fce,Projeto"
    procnom = "projeto"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Projetos"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "projeto=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(2).Text = rsLocal.Fields(2)
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub ChamaGridProduto()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "Select a.codigoprd,a.nomefantasia from " & vBancoTotvs & ".dbo.TPRD as a left join tbmateriais as b on a.IDPRD = b.idprd where a.CODIGOPRD like '%%' and a.codigoprd like '%" & txtcadastro(19) & "%' order by a.nomefantasia"
    procnom = "nomefantasia"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Materiais"
    Pesquisa = frmRelExpAvulso.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "CODIGOPRD=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If Pesquisa = "Lista de Materiais" Then Pesquisa = ""
            txtcadastro(19) = Pesquisa
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub CarregaDados(Index)
On Error GoTo Err
    Dim rsMaterial As New ADODB.Recordset
    SqlM = "Select a.CODIGOPRD,a.NOMEFANTASIA,b.formula,b.constpint,c.valconst,a.CODUNDCONTROLE,b.forpint,b.observacao,d.DESCRICAO,a.idprd from " & vBancoTotvs & ".dbo.tprd as a left join tbMateriais as b on b.idprd = a.idprd left Join tbconstantes as c on b.idprd = c.idprd left join " & vBancoTotvs & ".dbo.TTB2 as d on a.CODTB2FAT = d.CODTB2FAT where a.CODIGOPRD = '" & txtcadastro(19) & "'order by c.idseq"
    rsMaterial.Open SqlM, cnBanco, adOpenKeyset, adLockReadOnly
    Text1.Text = rsMaterial.Fields(1)
    rsMaterial.Close
    Set rsMaterial = Nothing
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

Private Sub txtCadastro_LostFocus(Index As Integer)
    voltaCorText txtcadastro(Index)
End Sub
