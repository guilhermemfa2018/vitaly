VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{76339437-30C4-11D4-AABA-0004ACBF1E11}#1.0#0"; "mcformresize.ocx"
Begin VB.Form CadProd 
   Caption         =   "CADASTRO DE PRODUTOS"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15015
   Icon            =   "CadProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   15015
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H8000000D&
      Height          =   9105
      Index           =   0
      Left            =   1920
      TabIndex        =   45
      Top             =   480
      Width           =   12855
      Begin VB.ComboBox Unidade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "CadProd.frx":08CA
         Left            =   240
         List            =   "CadProd.frx":08E9
         TabIndex        =   3
         Text            =   "UN"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Servi�o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   21
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Monitorar M�dia de Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9360
         TabIndex        =   23
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox ConsumoFim 
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
         Height          =   360
         Left            =   1680
         TabIndex        =   17
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Caracteristicas 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   240
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3840
         Width           =   6015
      End
      Begin VB.TextBox Periodo 
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
         Height          =   360
         Left            =   2640
         TabIndex        =   18
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox ConsumoIni 
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
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
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
         Left            =   8640
         TabIndex        =   47
         ToolTipText     =   "Inserir Localiza��o"
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
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
         Left            =   7560
         TabIndex        =   46
         ToolTipText     =   "Categorias"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Qt 
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
         Height          =   360
         Left            =   2040
         TabIndex        =   4
         Text            =   "1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox CodFabri 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2400
         TabIndex        =   0
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Fabricante 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox Descri�ao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8160
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.ComboBox Categoria 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "CadProd.frx":0911
         Left            =   4560
         List            =   "CadProd.frx":0913
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Referencia 
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
         Height          =   360
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Marca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Modelo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10680
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Ano 
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
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Especie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Tamanho 
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
         Height          =   360
         Left            =   9360
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Peso 
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
         Height          =   360
         Left            =   7680
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Localiza�ao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   19
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox Cor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Largura 
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
         Height          =   360
         Left            =   10920
         TabIndex        =   14
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Altura 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   6480
         TabIndex        =   15
         Top             =   3840
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "CadProd.frx":0915
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   8160
         OleObjectBlob   =   "CadProd.frx":0985
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "CadProd.frx":09F5
         TabIndex        =   51
         Top             =   240
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "CadProd.frx":0A75
         TabIndex        =   52
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "CadProd.frx":0AE7
         TabIndex        =   53
         Top             =   1080
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "CadProd.frx":0B53
         TabIndex        =   54
         Top             =   1080
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "CadProd.frx":0BC7
         TabIndex        =   55
         Top             =   1080
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   10680
         OleObjectBlob   =   "CadProd.frx":0C39
         TabIndex        =   56
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "CadProd.frx":0CA3
         TabIndex        =   57
         Top             =   1920
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "CadProd.frx":0D07
         TabIndex        =   58
         Top             =   1080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "CadProd.frx":0D6F
         TabIndex        =   59
         Top             =   2760
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "CadProd.frx":0DE3
         TabIndex        =   60
         Top             =   1920
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "CadProd.frx":0E49
         TabIndex        =   61
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "CadProd.frx":0EAD
         TabIndex        =   62
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   255
         Left            =   9360
         OleObjectBlob   =   "CadProd.frx":0F19
         TabIndex        =   63
         Top             =   1920
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "CadProd.frx":0F85
         TabIndex        =   64
         Top             =   1920
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   10920
         OleObjectBlob   =   "CadProd.frx":0FEB
         TabIndex        =   65
         Top             =   1920
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "CadProd.frx":105F
         TabIndex        =   68
         Top             =   2760
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "CadProd.frx":10D9
         TabIndex        =   69
         Top             =   2760
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel48 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "CadProd.frx":1153
         TabIndex        =   70
         Top             =   2760
         Width           =   615
      End
      Begin VB.Frame Frame9 
         Caption         =   "Forma��o de Pre�o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4335
         Left            =   120
         TabIndex        =   71
         Top             =   4680
         Width           =   12615
         Begin VB.CommandButton Command6 
            Caption         =   "&Aplicar �ndices  Padr�o"
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
            Left            =   8880
            TabIndex        =   117
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox chk_Padrao 
            Caption         =   "Gravar �ndices Como Padr�o"
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
            Left            =   7680
            TabIndex        =   116
            Top             =   3000
            Width           =   3615
         End
         Begin VB.CheckBox chk_Categoria 
            Caption         =   "Aplicar �ndices a Produtos da mesma Categoria"
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
            Left            =   960
            TabIndex        =   115
            Top             =   3600
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.CheckBox chk_Produtos 
            Caption         =   "Aplicar �ndices a Todos os Produtos"
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
            Left            =   5640
            TabIndex        =   114
            Top             =   3600
            Visible         =   0   'False
            Width           =   4575
         End
         Begin MSMask.MaskEdBox P�Compra 
            Height          =   420
            Left            =   960
            TabIndex        =   24
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   12648384
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   " "
         End
         Begin VB.TextBox FtLucro 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Left            =   3600
            TabIndex        =   25
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox IndIcms 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2640
            TabIndex        =   28
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IndCredIcms 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            TabIndex        =   27
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IndIpi 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4320
            TabIndex        =   29
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IndFrete 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6000
            TabIndex        =   30
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IndComissao 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   11040
            TabIndex        =   33
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IndCFixo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9360
            TabIndex        =   32
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox IndSimples 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7680
            TabIndex        =   31
            Top             =   1800
            Width           =   975
         End
         Begin MSMask.MaskEdBox P�Venda 
            Height          =   420
            Left            =   6000
            TabIndex        =   26
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   741
            _Version        =   393216
            BackColor       =   12648384
            ForeColor       =   12582912
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox P�Custo 
            Height          =   360
            Left            =   960
            TabIndex        =   34
            Top             =   3120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   32768
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox lucro 
            Height          =   360
            Left            =   3000
            TabIndex        =   35
            Top             =   3120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   192
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox LucroLiq 
            Height          =   360
            Left            =   5040
            TabIndex        =   36
            Top             =   3120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   192
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "CadProd.frx":11B7
            TabIndex        =   72
            Top             =   480
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
            Height          =   375
            Left            =   5520
            OleObjectBlob   =   "CadProd.frx":123F
            TabIndex        =   73
            Top             =   720
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
            Height          =   255
            Left            =   5040
            OleObjectBlob   =   "CadProd.frx":129F
            TabIndex        =   74
            Top             =   2880
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
            Height          =   255
            Left            =   3000
            OleObjectBlob   =   "CadProd.frx":1317
            TabIndex        =   75
            Top             =   2880
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "CadProd.frx":138B
            TabIndex        =   76
            Top             =   2880
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
            Height          =   255
            Left            =   6000
            OleObjectBlob   =   "CadProd.frx":1405
            TabIndex        =   77
            Top             =   480
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
            Height          =   255
            Left            =   3600
            OleObjectBlob   =   "CadProd.frx":147F
            TabIndex        =   78
            Top             =   480
            Width           =   1935
         End
         Begin MSMask.MaskEdBox ICMS 
            Height          =   360
            Left            =   2640
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox IPI 
            Height          =   360
            Left            =   4320
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FRETE 
            Height          =   360
            Left            =   6000
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Simples 
            Height          =   360
            Left            =   7680
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Fixo 
            Height          =   360
            Left            =   9360
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Comissao 
            Height          =   360
            Left            =   11040
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox CreditoIcms 
            Height          =   360
            Left            =   960
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$#,##0.00;(R$#,##0.00)"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   11040
            OleObjectBlob   =   "CadProd.frx":14FB
            TabIndex        =   86
            Top             =   1560
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   9360
            OleObjectBlob   =   "CadProd.frx":1569
            TabIndex        =   87
            Top             =   1560
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   7680
            OleObjectBlob   =   "CadProd.frx":15DB
            TabIndex        =   88
            Top             =   1560
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   6000
            OleObjectBlob   =   "CadProd.frx":1647
            TabIndex        =   89
            Top             =   1560
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   4320
            OleObjectBlob   =   "CadProd.frx":16AF
            TabIndex        =   90
            Top             =   1560
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "CadProd.frx":1713
            TabIndex        =   91
            Top             =   1560
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "CadProd.frx":177B
            TabIndex        =   92
            Top             =   1560
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   5400
            OleObjectBlob   =   "CadProd.frx":17EF
            TabIndex        =   93
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "CadProd.frx":184F
            TabIndex        =   94
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   12120
            OleObjectBlob   =   "CadProd.frx":18AF
            TabIndex        =   95
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   10440
            OleObjectBlob   =   "CadProd.frx":190F
            TabIndex        =   96
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   8760
            OleObjectBlob   =   "CadProd.frx":196F
            TabIndex        =   97
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   7080
            OleObjectBlob   =   "CadProd.frx":19CF
            TabIndex        =   98
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "CadProd.frx":1A2F
            TabIndex        =   99
            Top             =   1800
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "CadProd.frx":1A8F
            TabIndex        =   100
            Top             =   2355
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "CadProd.frx":1AFB
            TabIndex        =   101
            Top             =   1845
            Width           =   855
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enviar Para Balan�a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   22
         Top             =   3480
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "CadProd.frx":1B67
         TabIndex        =   118
         Top             =   3600
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "CadProd.frx":1BE3
         TabIndex        =   119
         Top             =   3600
         Width           =   2055
      End
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
      ForeColor       =   &H8000000D&
      Height          =   9135
      Index           =   1
      Left            =   1920
      TabIndex        =   66
      Top             =   480
      Width           =   12855
      Begin VB.Frame Frame3 
         Caption         =   "Classifica��o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4575
         Left            =   600
         TabIndex        =   108
         Top             =   4440
         Width           =   10935
         Begin VB.TextBox CodClassi 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7680
            TabIndex        =   111
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Excluir"
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
            Left            =   6360
            TabIndex        =   110
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            TabIndex        =   39
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Inserir"
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
            Left            =   5160
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H8000000A&
            Caption         =   "&�tens da Classifica��o"
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
            Left            =   3120
            TabIndex        =   41
            Top             =   4080
            Width           =   4335
         End
         Begin VB.CommandButton Command8 
            Appearance      =   0  'Flat
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
            Left            =   4440
            TabIndex        =   109
            ToolTipText     =   "Classificacar Categoria"
            Top             =   360
            Width           =   375
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "CadProd.frx":1C63
            Height          =   3135
            Left            =   240
            OleObjectBlob   =   "CadProd.frx":1C77
            TabIndex        =   112
            Top             =   840
            Width           =   10455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "CadProd.frx":2B6A
            TabIndex        =   113
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aplica��o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4095
         Left            =   600
         TabIndex        =   102
         Top             =   240
         Width           =   10935
         Begin VB.TextBox CodAplica�ao 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9000
            TabIndex        =   105
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Inserir"
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
            Left            =   5760
            TabIndex        =   38
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   37
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Excluir"
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
            Left            =   6960
            TabIndex        =   104
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
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
            Left            =   5160
            TabIndex        =   103
            ToolTipText     =   "Aplica��es"
            Top             =   360
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "CadProd.frx":2BD0
            TabIndex        =   106
            Top             =   360
            Width           =   1215
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "CadProd.frx":2C42
            Height          =   3015
            Left            =   240
            OleObjectBlob   =   "CadProd.frx":2C56
            TabIndex        =   107
            Top             =   840
            Width           =   10455
         End
      End
   End
   Begin prjcmcformresize.mcformresize mcformresize1 
      Left            =   600
      Top             =   4320
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   600
      Top             =   4920
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar (F9)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   42
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar (F11)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   44
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair (Esc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   43
      Top             =   2040
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   9615
      Left            =   1800
      TabIndex        =   67
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   16960
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PRODUTO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "APLICA��O E CLASSIFICA��O"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CadProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSair_Click()
Unload Me
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
    'LimpaTudo
ElseIf KeyCode = 123 Then 'F12
    'P�Compra.SetFocus
End If
End Sub

Private Sub Form_Load()
AplicarSkin Me, Principal.Skin1
NewColorDBGrid Me
On Error GoTo ErrHandler

Call BordasControle(Me, DBGrid1, False)
Call BordasControle(Me, DBGrid2, False)


AjustaContainer

Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte T�cnico.", , critico
End Sub
Private Sub AjustaContainer()
Dim i As Integer

With TabStrip1
    For i = 1 To .Tabs.Count
        Frame1(i - 1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
    Next
End With
Frame1(0).ZOrder
End Sub

Private Sub Timer2_Timer()
Me.WindowState = 2
End Sub
