VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes/Fornecedores"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5160
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   90
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   3120
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.CommandButton chamCad 
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
      Index           =   9
      Left            =   720
      Picture         =   "frmClientes.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   88
      Tag             =   "Sair"
      Top             =   7440
      Width           =   615
   End
   Begin VB.CommandButton chamCad 
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
      Left            =   120
      Picture         =   "frmClientes.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   89
      Tag             =   "Salvar"
      Top             =   7440
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      TabIndex        =   86
      Top             =   7440
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Tag             =   "Status do Cliente/Fornecedor"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ramo de Atividade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   60
      Top             =   120
      Width           =   7215
      Begin MSMask.MaskEdBox mskcadastro 
         Height          =   345
         Index           =   10
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do Ramo de Atividades do Cliente/Fornecedor"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmClientes.frx":265E
         TabIndex        =   62
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtcadastro 
         Height          =   345
         Index           =   23
         Left            =   1080
         TabIndex        =   1
         Tag             =   "Ramo de atividades do Cliente/Fornecedor"
         Top             =   480
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "frmClientes.frx":26C4
         TabIndex        =   61
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdcadastro 
         Caption         =   "..."
         Height          =   345
         Index           =   3
         Left            =   6720
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   43
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10610
      _Version        =   393216
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
      TabCaption(0)   =   "Jurídica"
      TabPicture(0)   =   "frmClientes.frx":2726
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Física"
      TabPicture(1)   =   "frmClientes.frx":2742
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contatos"
      TabPicture(2)   =   "frmClientes.frx":275E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Grid"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Dados "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   7
            Left            =   120
            TabIndex        =   16
            Tag             =   "Site do Cliente/Fornecedor"
            Top             =   4800
            Width           =   6735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   6
            Left            =   120
            TabIndex        =   15
            Tag             =   "Email do Cliente/Fornecedor"
            Top             =   4080
            Width           =   6735
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   3
            Left            =   5520
            TabIndex        =   14
            Tag             =   "Fax do Cliente/Fornecedor"
            Top             =   3360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   2
            Left            =   4080
            TabIndex        =   13
            Tag             =   "Telefone de contato do Cliente/Fornecedor"
            Top             =   3360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   1
            Left            =   2160
            TabIndex        =   12
            Tag             =   "Inscrição Estadual do Cliente/Fornecedor"
            Top             =   3360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###.######.####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Tag             =   "CNPJ do Cliente/Fornecedor"
            Top             =   3360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cbocadastro 
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
            ItemData        =   "frmClientes.frx":277A
            Left            =   6120
            List            =   "frmClientes.frx":27CF
            TabIndex        =   10
            Tag             =   "UF do Cliente/Fornecedor"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   5
            Left            =   3600
            TabIndex        =   9
            Tag             =   "Cidade do Cliente/Fornecedor"
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   4
            Left            =   1200
            TabIndex        =   8
            Tag             =   "Bairro do Cliente/Fornecedor"
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   21
            Left            =   120
            TabIndex        =   7
            Tag             =   "CEP do Cliente/Fornecedor"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Tag             =   "Endereço do Cliente/Fornecedor"
            Top             =   1920
            Width           =   6735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Tag             =   "Nome Fantasia do Cliente/Fornecedor"
            Top             =   1200
            Width           =   6735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   1
            Left            =   1080
            TabIndex        =   4
            Tag             =   "Razão Social do Cliente/Fornecedor"
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Tag             =   "Código do Cliente/Fornecedor"
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":283F
            TabIndex        =   76
            Top             =   4560
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":28A1
            TabIndex        =   75
            Top             =   3840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   345
            Left            =   5520
            OleObjectBlob   =   "frmClientes.frx":2905
            TabIndex        =   74
            Top             =   3120
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   345
            Left            =   4080
            OleObjectBlob   =   "frmClientes.frx":2965
            TabIndex        =   73
            Top             =   3120
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   345
            Left            =   2160
            OleObjectBlob   =   "frmClientes.frx":29CF
            TabIndex        =   72
            Top             =   3120
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2A45
            TabIndex        =   71
            Top             =   3120
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   345
            Left            =   6120
            OleObjectBlob   =   "frmClientes.frx":2AA7
            TabIndex        =   70
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   345
            Left            =   3600
            OleObjectBlob   =   "frmClientes.frx":2B0D
            TabIndex        =   69
            Top             =   2400
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   345
            Left            =   1200
            OleObjectBlob   =   "frmClientes.frx":2B73
            TabIndex        =   68
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2BD9
            TabIndex        =   67
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2C39
            TabIndex        =   66
            Top             =   1680
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2CA3
            TabIndex        =   65
            Top             =   960
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2D17
            TabIndex        =   63
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Razão Social:"
            Height          =   345
            Left            =   1080
            TabIndex        =   64
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   6975
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   6
            Left            =   1920
            TabIndex        =   40
            Tag             =   "Celular do Cliente/Fornecedor"
            Top             =   2640
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)#####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   5
            Left            =   120
            TabIndex        =   39
            Tag             =   "Telefone do Cliente/Fornecedor"
            Top             =   2640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   4
            Left            =   5280
            TabIndex        =   38
            Tag             =   "CPF do Cliente/Fornecedor"
            Top             =   2640
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   120
            TabIndex        =   30
            Tag             =   "Código de Identificação do Cliente/Fornecedor"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   16
            Left            =   120
            TabIndex        =   42
            Tag             =   "Site do Cliente/Fornecedor"
            Top             =   4080
            Width           =   6735
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   15
            Left            =   120
            TabIndex        =   41
            Tag             =   "Endereco de email do Cliente/Fornecedor"
            Top             =   3360
            Width           =   6735
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   14
            Left            =   3600
            TabIndex        =   37
            Tag             =   "RG do Cliente/Fornecedor"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.ComboBox cbocadastro 
            Height          =   345
            Index           =   1
            ItemData        =   "frmClientes.frx":2D7D
            Left            =   6000
            List            =   "frmClientes.frx":2DD2
            TabIndex        =   36
            Tag             =   "UF do Cliente/Fornecedor"
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   13
            Left            =   4080
            TabIndex        =   35
            Tag             =   "Cidade do Cliente/Fornecedor"
            Top             =   1920
            Width           =   2775
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   12
            Left            =   1200
            TabIndex        =   34
            Tag             =   "Bairro do Cliente/Fornecedor"
            Top             =   1920
            Width           =   2775
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   11
            Left            =   120
            TabIndex        =   33
            Tag             =   "CEP do Cliente/Fornecedor"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   10
            Left            =   120
            TabIndex        =   32
            Tag             =   "Endereço do Cliente/Fornecedor"
            Top             =   1200
            Width           =   5775
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   9
            Left            =   1080
            TabIndex        =   31
            Tag             =   "Nome do Cliente/Fornecedor"
            Top             =   480
            Width           =   5775
         End
         Begin VB.Label Label21 
            Caption         =   "Código:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Site:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Email:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Celular:"
            Height          =   255
            Left            =   1920
            TabIndex        =   55
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Telefone:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "CPF:"
            Height          =   255
            Left            =   5280
            TabIndex        =   53
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Identidade:"
            Height          =   255
            Left            =   3600
            TabIndex        =   52
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "UF:"
            Height          =   255
            Left            =   6000
            TabIndex        =   51
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade:"
            Height          =   255
            Left            =   4080
            TabIndex        =   50
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Bairro:"
            Height          =   255
            Left            =   1200
            TabIndex        =   49
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "CEP:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Endereço:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Nome:"
            Height          =   255
            Left            =   1080
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   29
         Top             =   4200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   -2147483634
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
      Begin VB.Frame Frame1 
         Caption         =   "Dados "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   44
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   22
            Left            =   120
            TabIndex        =   25
            Tag             =   "Endereço de email do contato"
            Top             =   2640
            Width           =   6735
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   9
            Left            =   4920
            TabIndex        =   24
            Tag             =   "Celular"
            Top             =   1920
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)#####-####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   24
            Left            =   4080
            TabIndex        =   23
            Tag             =   "Ramal"
            Top             =   1920
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   8
            Left            =   2040
            TabIndex        =   22
            Tag             =   "Fax"
            Top             =   1920
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskcadastro 
            Height          =   345
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Tag             =   "Telefone de contato"
            Top             =   1920
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(##)####-####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   20
            Left            =   3720
            TabIndex        =   20
            Tag             =   "Função ocupada"
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   19
            Left            =   120
            TabIndex        =   19
            Tag             =   "Cargo ocupado"
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   18
            Left            =   3720
            TabIndex        =   18
            Tag             =   "Nome do departamento"
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   17
            Left            =   120
            TabIndex        =   17
            Tag             =   "Nome do Contato"
            Top             =   480
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2E42
            TabIndex        =   85
            Top             =   2400
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   4920
            OleObjectBlob   =   "frmClientes.frx":2EA6
            TabIndex        =   84
            Top             =   1680
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "frmClientes.frx":2F0E
            TabIndex        =   83
            Top             =   1680
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "frmClientes.frx":2F72
            TabIndex        =   82
            Top             =   1680
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":2FD2
            TabIndex        =   81
            Top             =   1680
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmClientes.frx":3034
            TabIndex        =   80
            Top             =   960
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmClientes.frx":309A
            TabIndex        =   79
            Top             =   240
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":310C
            TabIndex        =   78
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmClientes.frx":316E
            TabIndex        =   77
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdcadastro 
            Caption         =   "&Excluir"
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   28
            Tag             =   "Excluir contato selecionado"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdcadastro 
            Caption         =   "&Alterar"
            Height          =   495
            Index           =   1
            Left            =   1440
            TabIndex        =   27
            Tag             =   "Alterar dados do contato selecionado"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdcadastro 
            Caption         =   "&Incluir"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Tag             =   "Incluir contato na lista abaixo"
            Top             =   3120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsCliForJ As New ADODB.Recordset
Private rsCliForF As New ADODB.Recordset
Private rsCliFor As New ADODB.Recordset
Private rsLocal As New ADODB.Recordset
Private rsLocal1 As New ADODB.Recordset
    
Private SqlJ As String
Private SqlF As String
Private SqlM As String
Private SqlLocal1 As String

Private Sqlpj As String
Private Sqlpf As String

Private ByLinhaInclusaoGrid As Integer
Private smensagem As String
Private Binclusao As Boolean
Private TipoCad As String
Private Status As String

Private Sub chamCad_Click(Index As Integer)
    Select Case Index
    Case 8
        If ValidaCampo = False Then Exit Sub
        'CancelaSN = 1
        Bot_salvar
        AtualizaListview
        Unload Me
        Set frmClientes = Nothing
    Case 9
        If Msgbox("Deseja sair da tela de cadastro?", vbQuestion + vbYesNo, "Zeus") = vbYes Then
            'CancelaSN = 0
            Unload Me
            Set frmClientes = Nothing
        End If
    End Select
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
On Error GoTo Err
    Dim conteudo As String
    Select Case Index
    Case 0
        IncluirItem
        If Me.Grid.Rows > 1 Then
            cmdCadastro(1).Enabled = True
            cmdCadastro(2).Enabled = True
        End If
    Case 1
        AlterarItem
    Case 2
        ExcluirItem
    Case 3
        Dim F As New frmpesqger
        Dim Iposicao As Variant
        Sqlp = "Select * from tbAtividades order by descricao"
        procnom = "descricao"
        campo = 1
        Campo1 = 0
        Load F
        F.Caption = "Pesquisa de Ramo de Atividade"
        Pesquisa = frmClientes.Tag
        F.Show 1
        If Pesquisa <> "" Then
            rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
            If rsLocal.RecordCount < 1 Then Exit Sub
            rsLocal.MoveFirst
            rsLocal.Find "descricao=" & "'" & Pesquisa & "'"
            If Not rsLocal.EOF Then
                mskCadastro(10).Text = Format(rsLocal.Fields(0), "000000")
                txtcadastro(23).Text = rsLocal.Fields(1)
            Else
                Msgbox "Ramo de Atividade não cadastrado", vbInformation, "Zeus"
            End If
            rsLocal.Close
            Set rsLocal = Nothing
        End If
    End Select
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    inicializa_tabs SSTab1, Picture1
    AbrirClientes
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabEnabled(2) = True
    SSTab1.Tab = 0
    If rsCliFor.RecordCount > 0 Then
        rsCliFor.MoveLast
        CompoeControles
    Else
        LimpaControles
    End If
    SSTab1.Enabled = True
    TipoCad = Pesquisa
    If TipoCad = "novo" Then
        LimpaControles
    ElseIf TipoCad = "editar" Then
        ResultPesq
        DesbloqueiaControles
    End If
    FecharClientes
    carregarIconBotao
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao chamCad(8), 8, 45 'Salvar
    carregaImagemBotao chamCad(9), 9, 34 'Sair
End Sub

Private Sub AbrirClientes()
On Error GoTo Err
    SqlM = "Select * from tbcliFor Order by codclifor"
    'Sqlp = SqlM
    rsCliFor.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    
    SqlJ = "Select * from tbcliFor, tbjuridica,tbatividades where tbjuridica.codclifor = tbclifor.codclifor and tbclifor.codatividade = tbatividades.codigo order by tbclifor.codclifor"
    Sqlpj = SqlJ
    rsCliForJ.Open SqlJ, cnBanco, adOpenKeyset, adLockOptimistic
    
    SqlF = "Select * from tbcliFor, tbfisica,tbatividades where tbfisica.codclifor = tbclifor.codclifor and tbclifor.codatividade = tbatividades.codigo order by tbclifor.codclifor"
    Sqlpf = SqlF
    rsCliForF.Open SqlF, cnBanco, adOpenKeyset, adLockOptimistic
    
    SqlLocal1 = "Select * from tbAtividades order by descricao"
    rsLocal1.Open SqlLocal1, cnBanco, adOpenKeyset, adLockOptimistic
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
Private Sub FecharClientes()
    rsCliFor.Close
    Set rsCliFor = Nothing
    rsCliForF.Close
    Set rsCliForF = Nothing
    rsCliForJ.Close
    Set rsCliForJ = Nothing
    rsLocal1.Close
    Set rsLocal1 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub mskCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Err
    If KeyCode = 13 Then
        Dim SqlLocal As String
        SqlLocal = "Select * from tbAtividades where tbAtividades.codigo = '" & Val(Me.mskCadastro(10)) & "'"
        rsLocal.Open SqlLocal, cnBanco, adOpenKeyset, adLockOptimistic
        
        If rsLocal.RecordCount = 0 Then
            mskCadastro(10).PromptInclude = False
            mskCadastro(10).Text = Format(mskCadastro(10), "000000") & ""
            mskCadastro(10).PromptInclude = True
            Msgbox "Código não cadastrado"
            mskCadastro(10).SetFocus
        Else
            mskCadastro(10).PromptInclude = False
            mskCadastro(10).Text = Format(rsLocal.Fields(0), "000000") & ""
            mskCadastro(10).PromptInclude = True
            txtcadastro(23).Text = rsLocal.Fields(1)
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

Private Sub optCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        SSTab1.Enabled = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 0
    Case 1
        SSTab1.Enabled = True
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 1
    End Select
End Sub

Private Sub LimpaControles()
    Dim x As Integer
    DesbloqueiaControles
    For x = 0 To mskCadastro.Count - 1
        mskCadastro(x).PromptInclude = False
        mskCadastro(x) = ""
        mskCadastro(x).PromptInclude = True
    Next
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x) = ""
    Next
    For x = 0 To cboCadastro.Count - 1
        cboCadastro(x) = ""
    Next

    Grid.Rows = 2
    Grid.Cols = 12
    Me.Grid.ColWidth(0) = 200
    Me.Grid.ColWidth(1) = 0
    Me.Grid.ColWidth(2) = 3000
    Me.Grid.ColAlignment(2) = flexAlignLeftCenter
    Me.Grid.ColWidth(3) = 1500
    Me.Grid.ColAlignment(3) = flexAlignLeftCenter
    
    Me.Grid.ColWidth(4) = 1500
    Me.Grid.ColAlignment(4) = flexAlignLeftCenter
    Me.Grid.ColWidth(5) = 1500
    Me.Grid.ColAlignment(5) = flexAlignLeftCenter
    Me.Grid.ColWidth(6) = 1500
    Me.Grid.ColAlignment(6) = flexAlignLeftCenter
    Me.Grid.ColWidth(7) = 1500
    Me.Grid.ColAlignment(7) = flexAlignLeftCenter
    Me.Grid.ColWidth(8) = 1500
    Me.Grid.ColAlignment(8) = flexAlignLeftCenter
    Me.Grid.ColWidth(9) = 4000
    Me.Grid.ColAlignment(9) = flexAlignLeftCenter
    Me.Grid.ColWidth(10) = 4000
    Me.Grid.ColAlignment(10) = flexAlignLeftCenter
        
    Me.Grid.TextMatrix(0, 2) = "Nome"
    Me.Grid.TextMatrix(0, 3) = "Departamento"
    Me.Grid.TextMatrix(0, 4) = "Cargo"
    Me.Grid.TextMatrix(0, 5) = "Função"
    Me.Grid.TextMatrix(0, 6) = "Fone"
    Me.Grid.TextMatrix(0, 7) = "Ramal"
    Me.Grid.TextMatrix(0, 8) = "Fax"
    Me.Grid.TextMatrix(0, 9) = "Celular"
    Me.Grid.TextMatrix(0, 10) = "Email"
    Me.Grid.TextMatrix(0, 11) = "Ligação"
    
    Binclusao = True
    Me.Grid.Rows = Me.Grid.FixedRows
    Me.Grid.Rows = Me.Grid.FixedRows + 1
    
    'chkCadastro(1).SetFocus
        
End Sub
Private Sub LimpaControleItem()
    Me.txtcadastro(17).Text = ""
    Me.txtcadastro(18).Text = ""
    Me.txtcadastro(19).Text = ""
    Me.txtcadastro(20).Text = ""
    Me.txtcadastro(22).Text = ""
    Me.txtcadastro(24).Text = ""
    Me.mskCadastro(7).PromptInclude = False
    Me.mskCadastro(7).Text = ""
    Me.mskCadastro(7).PromptInclude = True
    Me.mskCadastro(8).PromptInclude = False
    Me.mskCadastro(8).Text = ""
    Me.mskCadastro(8).PromptInclude = True
    Me.mskCadastro(9).PromptInclude = False
    Me.mskCadastro(9).Text = ""
    Me.mskCadastro(9).PromptInclude = True
    'Me.cboCadastro(4).Text = ""
End Sub

Private Sub CompoeControles()
    Dim Z As Integer
   
    Dim SqlJ As String
    Dim SqlF As String
    Dim x As Integer
    BloqueiaControles
    If SSTab1.TabEnabled(0) = True Then
        If rsCliForJ.RecordCount > 0 Then
            txtcadastro(0).Text = Format(rsCliForJ.Fields(0), "000000") & ""
            mskCadastro(0).PromptInclude = False
            mskCadastro(0).Text = rsCliForJ.Fields(19) & ""
            mskCadastro(0).PromptInclude = True
            mskCadastro(1).PromptInclude = False
            mskCadastro(1).Text = rsCliForJ.Fields(20) & ""
            mskCadastro(1).PromptInclude = True
            mskCadastro(2).PromptInclude = False
            mskCadastro(2).Text = rsCliForJ.Fields(6) & ""
            mskCadastro(2).PromptInclude = True
            mskCadastro(3).PromptInclude = False
            mskCadastro(3).Text = rsCliForJ.Fields(7) & ""
            mskCadastro(3).PromptInclude = True
            mskCadastro(10).PromptInclude = False
            mskCadastro(10).Text = Format(rsCliForJ.Fields(12), "000000") & ""
            mskCadastro(10).PromptInclude = True
            txtcadastro(3).Text = rsCliForJ.Fields(1) & ""
            txtcadastro(21).Text = rsCliForJ.Fields(2) & ""
            txtcadastro(4).Text = rsCliForJ.Fields(3) & ""
            txtcadastro(5).Text = rsCliForJ.Fields(4) & ""
            txtcadastro(6).Text = rsCliForJ.Fields(8) & ""
            txtcadastro(7).Text = rsCliForJ.Fields(9) & ""
            txtcadastro(1).Text = rsCliForJ.Fields(17) & ""
            txtcadastro(2).Text = rsCliForJ.Fields(18) & ""
            txtcadastro(23) = rsCliForJ.Fields(22) & ""
            cboCadastro(0).Text = rsCliForJ.Fields(5) & ""
            
            If rsCliForJ.Fields(15) = "S" Then
                Check1.Value = 1
            Else
                Check1.Value = 0
            End If
            
            'optCadastro(0).Value = True
            
        End If
    'ElseIf SSTab1.TabEnabled(0) = False Then
    '    If rsCliForF.RecordCount > 0 Then
    '        txtcadastro(8).Text = Format(rsCliForF.Fields(0), "000000") & ""
    '        mskcadastro(4).PromptInclude = False
    '        mskcadastro(4).Text = rsCliForF.Fields(18) & ""
    '        mskcadastro(4).PromptInclude = True
    '        mskcadastro(5).PromptInclude = False
    '        mskcadastro(5).Text = rsCliForF.Fields(6) & ""
    '        mskcadastro(5).PromptInclude = True
    '        mskcadastro(6).PromptInclude = False
    '        mskcadastro(6).Text = rsCliForF.Fields(7) & ""
    '        mskcadastro(6).PromptInclude = True
    '
    '        mskcadastro(10).PromptInclude = False
    '        mskcadastro(10).Text = Format(rsCliForF.Fields(19), "000000") & ""
    '        mskcadastro(10).PromptInclude = True
    '
    '        txtcadastro(9).Text = rsCliForF.Fields(16) & ""
    '        txtcadastro(10).Text = rsCliForF.Fields(1) & ""
    '        txtcadastro(11).Text = rsCliForF.Fields(2) & ""
    '        txtcadastro(12).Text = rsCliForF.Fields(3) & ""
    '        txtcadastro(13).Text = rsCliForF.Fields(4) & ""
    '        txtcadastro(14).Text = rsCliForF.Fields(17) & ""
    '        txtcadastro(15).Text = rsCliForF.Fields(8) & ""
    '        txtcadastro(16).Text = rsCliForF.Fields(9) & ""
    '        txtcadastro(23) = rsCliForF.Fields(20) & ""
    '        cbocadastro(1).Text = rsCliForF.Fields(5) & ""
    '
    '        'optCadastro(1).Value = True
    '
    '    End If
    End If
    CompoeGrid
End Sub
Private Sub CompoeGrid()
On Error GoTo Err
    Dim rsGrid As New ADODB.Recordset
    Dim x As Integer, y As Integer
    Dim Soma As Integer
    Dim SqlGrid As String
    Dim CTotal As Currency
    Grid.Rows = Grid.FixedRows ' nº de linha da grade
    Grid.Rows = Grid.FixedRows + 1
    If txtcadastro(1) <> "" Or txtcadastro(8) <> "" Then
        If SSTab1.TabEnabled(1) = False Then
            SqlGrid = "select tbcontatos.nome, tbcontatos.departamento, tbcontatos.cargo, tbcontatos.funcao, tbcontatos.telefone, tbcontatos.fax, tbcontatos.celular, tbcontatos.email, tbcontatos.ramal, tbcontatos.tipolig from tbclifor, tbcontatos where tbcontatos.codclifor = '" & Val(Me.txtcadastro(0)) & "'" & _
            "and tbclifor.codclifor = '" & Val(Me.txtcadastro(0)) & "'"
        ElseIf SSTab1.TabEnabled(0) = False Then
            SqlGrid = "select tbcontatos.nome, tbcontatos.departamento, tbcontatos.cargo, tbcontatos.funcao, tbcontatos.telefone, tbcontatos.fax, tbcontatos.celular, tbcontatos.email, tbcontatos.ramal, tbcontatos.tipolig from tbclifor, tbcontatos where tbcontatos.codclifor = '" & Val(Me.txtcadastro(8)) & "'" & _
            "and tbclifor.codclifor = '" & Val(Me.txtcadastro(8)) & "'"
        End If
    End If
    rsGrid.Open SqlGrid, cnBanco, adOpenKeyset, adLockOptimistic
    
    Grid.Rows = Grid.Rows + rsGrid.RecordCount
    Grid.Cols = 12
    Me.Grid.ColWidth(0) = 200
    Me.Grid.ColWidth(1) = 0
    Me.Grid.ColWidth(2) = 3000
    Me.Grid.ColAlignment(2) = flexAlignLeftCenter
    Me.Grid.ColWidth(3) = 1500
    Me.Grid.ColAlignment(3) = flexAlignLeftCenter
    Me.Grid.ColWidth(4) = 1500
    Me.Grid.ColAlignment(4) = flexAlignLeftCenter
    Me.Grid.ColWidth(5) = 1500
    Me.Grid.ColAlignment(5) = flexAlignLeftCenter
    Me.Grid.ColWidth(6) = 1500
    Me.Grid.ColAlignment(6) = flexAlignLeftCenter
    Me.Grid.ColWidth(7) = 1500
    Me.Grid.ColAlignment(7) = flexAlignLeftCenter
    Me.Grid.ColWidth(8) = 1500
    Me.Grid.ColAlignment(8) = flexAlignLeftCenter
    Me.Grid.ColWidth(9) = 4000
    Me.Grid.ColAlignment(9) = flexAlignLeftCenter
    Me.Grid.ColWidth(10) = 4000
    Me.Grid.ColAlignment(10) = flexAlignLeftCenter
    Me.Grid.ColWidth(11) = 4000
    Me.Grid.ColAlignment(11) = flexAlignLeftCenter
        
    Me.Grid.TextMatrix(0, 2) = "Nome"
    Me.Grid.TextMatrix(0, 3) = "Departamento"
    Me.Grid.TextMatrix(0, 4) = "Cargo"
    Me.Grid.TextMatrix(0, 5) = "Função"
    Me.Grid.TextMatrix(0, 6) = "Fone"
    Me.Grid.TextMatrix(0, 7) = "Fax"
    Me.Grid.TextMatrix(0, 8) = "Celular"
    Me.Grid.TextMatrix(0, 9) = "Email"
    Me.Grid.TextMatrix(0, 10) = "Ramal"
    Me.Grid.TextMatrix(0, 11) = "Ligação"
    
    If rsGrid.RecordCount > 0 Then
        Do While Not rsGrid.EOF
            Soma = Soma + 1
            rsGrid.MoveNext
        Loop
        rsGrid.MoveFirst
        For x = 1 To Soma
            Me.Grid.Row = x
            For y = 1 To rsGrid.Fields.Count
                Me.Grid.Col = y + 1
                If y > 1 Then
                     Me.Grid.Text = Format(rsGrid.Fields(y - 1), "##0")
                Else
                    Me.Grid.Text = rsGrid.Fields(y - 1)
                End If
            Next
            rsGrid.MoveNext
        Next
    End If
    rsGrid.Close
    Set rsGrid = Nothing
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

Private Sub IncluirItem()
    Dim x As Integer
    Dim CTotal As Currency
    
    If ValidaItem(smensagem) Then
        
        If ByLinhaInclusaoGrid = 0 Then
            ByLinhaInclusaoGrid = Me.Grid.Rows - 1
            Me.Grid.Rows = Me.Grid.Rows + 1
        End If
        CTotal = 0
        Me.Grid.ColAlignment(2) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 2) = Me.txtcadastro(17).Text
        
        Me.Grid.ColAlignment(3) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 3) = Me.txtcadastro(18).Text
        
        Me.Grid.ColAlignment(4) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 4) = Me.txtcadastro(19).Text
        
        Me.Grid.ColAlignment(5) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 5) = Me.txtcadastro(20).Text
        
        Me.Grid.ColAlignment(9) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 9) = Me.txtcadastro(22).Text
               
        Me.Grid.ColAlignment(6) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 6) = RemoveMask(mskCadastro(7))
        
        Me.Grid.ColAlignment(7) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 7) = RemoveMask(mskCadastro(8))
                      
        Me.Grid.ColAlignment(8) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 8) = RemoveMask(mskCadastro(9))
       
        Me.Grid.ColAlignment(10) = flexAlignLeftCenter
        Me.Grid.TextMatrix(ByLinhaInclusaoGrid, 10) = Me.txtcadastro(24).Text
        
        Me.Grid.ColAlignment(11) = flexAlignLeftCenter
        LimpaControleItem
        ByLinhaInclusaoGrid = 0
    Else
        Msgbox smensagem, vbInformation, "ATENÇÃO"
    End If
End Sub

Private Sub AlterarItem()
    If Me.Grid.RowSel <> 0 Then
        Me.txtcadastro(17).Text = Me.Grid.TextMatrix(Me.Grid.RowSel, 2)
        Me.txtcadastro(18).Text = Me.Grid.TextMatrix(Me.Grid.RowSel, 3)
        Me.txtcadastro(19).Text = Me.Grid.TextMatrix(Me.Grid.RowSel, 4)
        Me.txtcadastro(20).Text = Me.Grid.TextMatrix(Me.Grid.RowSel, 5)
        Me.txtcadastro(22).Text = Me.Grid.TextMatrix(Me.Grid.RowSel, 9)
        Me.txtcadastro(24).Text = Me.Grid.TextMatrix(Me.Grid.RowSel, 10)
        If Me.Grid.TextMatrix(Me.Grid.RowSel, 6) <> "" Then Me.mskCadastro(7).Text = Format(Me.Grid.TextMatrix(Me.Grid.RowSel, 6), "(##)####-####")
        If Me.Grid.TextMatrix(Me.Grid.RowSel, 7) <> "" Then Me.mskCadastro(8).Text = Format(Me.Grid.TextMatrix(Me.Grid.RowSel, 7), "(##)####-####")
        If Me.Grid.TextMatrix(Me.Grid.RowSel, 8) <> "" Then Me.mskCadastro(9).Text = Format(Me.Grid.TextMatrix(Me.Grid.RowSel, 8), "(##)####-####")
        ByLinhaInclusaoGrid = Me.Grid.RowSel
    End If
End Sub

Private Sub ExcluirItem()
    Dim x As Integer
    Dim VetorGrid() As Variant
    Dim ByLinhaSelecionada As Byte
    Dim lGrid, lVetor, y As Byte
    If Me.Grid.RowSel <> Me.Grid.Rows - 1 Then
        ReDim VetorGrid(Me.Grid.Rows - 1, 10)
        ByLinhaSelecionada = Me.Grid.RowSel
        For lGrid = 0 To Me.Grid.Rows - 1
            For y = 1 To 10
                VetorGrid(lGrid, y) = Me.Grid.TextMatrix(lGrid, y)
            Next
        Next
        Me.Grid.Rows = Me.Grid.Rows - 1
        lGrid = 0
        For lVetor = 0 To UBound(VetorGrid())
            If lVetor <> ByLinhaSelecionada Then
                For y = 1 To 10
                    Me.Grid.TextMatrix(lGrid, y) = VetorGrid(lVetor, y)
                Next
                lGrid = lGrid + 1
            End If
        Next
        ByLinhaInclusaoGrid = 0
    End If
    Erase VetorGrid
End Sub

Private Sub Grid_DblClick()
    AlterarItem
End Sub

Private Sub Bot_salvar()
On Error GoTo TrataErro
    If Msgbox("Confirma o cadastramento dos dados?", vbQuestion + vbYesNo, "Atenção") = vbNo Then Exit Sub
    If ValidaCampo = False Then Exit Sub
    Dim SqlM As String
    Dim SqlGpj As String
    Dim SqlGpf As String
    Dim SqlGrid As String
    Dim x, y, CodLV As Integer
    Dim rsGrid As New ADODB.Recordset
    Dim rsGpf As New ADODB.Recordset
    Dim rsGpj As New ADODB.Recordset
    Dim rsGCF As New ADODB.Recordset
10  cnBanco.BeginTrans ' Inicia a transação
    If SSTab1.TabEnabled(0) = True Then
        If txtcadastro(1).Text <> "" Then
            SqlM = "Select * from tbclifor where tbclifor.codclifor= " & Val(Me.txtcadastro(0))
            rsGCF.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
            CodLV = 0
            If txtcadastro(0).Text = "" Then
                rsGCF.AddNew
                CodLV = GeraCodigo
                rsGCF.Fields(0) = CodLV
            End If
        
            rsGCF.Fields(6) = RemoveMask(mskCadastro(2).ClipText)
            rsGCF.Fields(7) = RemoveMask(mskCadastro(3).ClipText)
            rsGCF.Fields(12) = mskCadastro(10)
            rsGCF.Fields(1) = txtcadastro(3).Text
            rsGCF.Fields(2) = txtcadastro(21).Text
            rsGCF.Fields(3) = txtcadastro(4).Text
            rsGCF.Fields(4) = txtcadastro(5).Text
            rsGCF.Fields(8) = txtcadastro(6).Text
            rsGCF.Fields(9) = txtcadastro(7).Text
            rsGCF.Fields(13) = txtcadastro(2).Text
            rsGCF.Fields(5) = cboCadastro(0).Text
            rsGCF(11) = 1
            rsGCF(10) = 1
            rsGCF.Fields(15) = "S"
            
            rsGCF.Update
            
            rsGCF.Close
            Set rsGCF = Nothing
            
            
            SqlGpj = "Select * from tbjuridica where tbjuridica.codclifor= " & Val(Me.txtcadastro(0))
            rsGpj.Open SqlGpj, cnBanco, adOpenKeyset, adLockOptimistic

            If txtcadastro(0).Text = "" Then
                rsGpj.AddNew
                rsGpj.Fields(0) = GeraCodigo - 1
            End If
            rsGpj.Fields(3) = mskCadastro(0).ClipText
            rsGpj.Fields(4) = mskCadastro(1).ClipText
            rsGpj.Fields(1) = txtcadastro(1).Text
            rsGpj.Fields(2) = txtcadastro(2).Text
            rsGpj.Update
    
            SqlGrid = "Delete from tbcontatos where tbcontatos.codclifor= " & Val(Me.txtcadastro(0))
            rsGrid.Open SqlGrid, cnBanco
        
            SqlGrid = "Select * from tbcontatos where tbcontatos.codclifor= " & Val(Me.txtcadastro(0))
            rsGrid.Open SqlGrid, cnBanco, adOpenKeyset, adLockOptimistic

            If rsGrid.RecordCount > 1 Then rsGrid.MoveLast
            y = 0
            With rsGrid
                For x = 1 To Me.Grid.Rows - 2
                    If Me.Grid.TextMatrix(x, 2) <> "" Then
                        y = y + 1
                        .AddNew
                        If txtcadastro(0) = "" Then
                            .Fields(0) = GeraCodigo - 1
                        Else
                            .Fields(0) = txtcadastro(0)
                        End If
                        .Fields(1) = y
                        .Fields(2) = Me.Grid.TextMatrix(x, 2)
                        .Fields(3) = Me.Grid.TextMatrix(x, 3)
                        .Fields(4) = Me.Grid.TextMatrix(x, 4)
                        .Fields(5) = Me.Grid.TextMatrix(x, 5)
                        .Fields(6) = Me.Grid.TextMatrix(x, 6)
                        .Fields(7) = Me.Grid.TextMatrix(x, 7)
                        .Fields(8) = Me.Grid.TextMatrix(x, 8)
                        .Fields(9) = Me.Grid.TextMatrix(x, 9)
                        .Fields(10) = Me.Grid.TextMatrix(x, 10)
                        .Fields(11) = Val(Me.Grid.TextMatrix(x, 11))
                        .Update
                    End If
                Next
            End With
            SSTab1.Tab = 0
            txtcadastro(1).SetFocus
        Else
            Msgbox "Favor Preencher o campo!", vbInformation, "ZEUS"
        End If
        rsGpj.Close
        Set rsGpj = Nothing
    End If
    'If SSTab1.TabEnabled(0) = False Then
    '    If txtcadastro(9).Text <> "" Then
    '        SqlM = "Select * from tbclifor where tbclifor.codclifor= " & Val(Me.txtcadastro(8))
    '        rsGCF.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    '        CodLV = 0
    '
    '        If txtcadastro(8).Text = "" Then
    '            rsGCF.AddNew
    '            CodLV = GeraCodigo
    '            rsGCF.Fields(0) = CodLV
     '       End If
    '
    '        rsGCF.Fields(6) = RemoveMask(mskcadastro(5).ClipText)
     '       rsGCF.Fields(7) = RemoveMask(mskcadastro(6).ClipText)
    '        rsGCF.Fields(12) = mskcadastro(10).ClipText
    '        rsGCF.Fields(1) = txtcadastro(10).Text
    '        rsGCF.Fields(2) = txtcadastro(11).Text
    '        rsGCF.Fields(3) = txtcadastro(12).Text
    '        rsGCF.Fields(4) = txtcadastro(13).Text
    '        rsGCF.Fields(8) = txtcadastro(15).Text
    '        rsGCF.Fields(9) = txtcadastro(16).Text
    '        rsGCF.Fields(13) = txtcadastro(9).Text
    '        rsGCF.Fields(5) = cbocadastro(1).Text
    '        rsGCF(11) = 1
    '        rsGCF(10) = 2
    '        rsGCF.Update
    '
    '        SqlGpf = "Select * from tbfisica where tbfisica.codclifor= " & Val(Me.txtcadastro(8))
    '        rsGpf.Open SqlGpf, cnBanco, adOpenKeyset, adLockOptimistic
    '
    '        If txtcadastro(8).Text = "" Then
    '            rsGpf.AddNew
    '            rsGpf.Fields(0) = GeraCodigo - 1
    '        End If
    '        rsGpf.Fields(3) = mskcadastro(4).ClipText
    '        rsGpf.Fields(1) = txtcadastro(9).Text
    '        rsGpf.Fields(2) = txtcadastro(14).Text
    '        rsGpf.Update
   '
   '         SqlGrid = "Delete from tbcontatos where tbcontatos.codclifor= " & Val(Me.txtcadastro(8))
   '         rsGrid.Open SqlGrid, cnBanco
   '
   '         SqlGrid = "Select * from tbcontatos where tbcontatos.codclifor= " & Val(Me.txtcadastro(8))
   '         rsGrid.Open SqlGrid, cnBanco, adOpenKeyset, adLockOptimistic
   '
   '         If rsGrid.RecordCount > 1 Then rsGrid.MoveLast
   '         Y = 0
   '         With rsGrid
   '             For X = 1 To Me.Grid.Rows - 2
   '                 If Me.Grid.TextMatrix(X, 2) <> "" Then
   '                     Y = Y + 1
   '                     .AddNew
   '                     If txtcadastro(0) = "" Then
   '                         .Fields(0) = GeraCodigo - 1
   '                     Else
   '                         .Fields(0) = txtcadastro(8)
   '                     End If
   '                     .Fields(1) = Y
   '                     .Fields(2) = Me.Grid.TextMatrix(X, 2)
   '                     .Fields(3) = Me.Grid.TextMatrix(X, 3)
   '                     .Fields(4) = Me.Grid.TextMatrix(X, 4)
   '                     .Fields(5) = Me.Grid.TextMatrix(X, 5)
   '                     .Fields(6) = Me.Grid.TextMatrix(X, 6)
   '                     .Fields(7) = Me.Grid.TextMatrix(X, 8)
   '                     .Fields(8) = Me.Grid.TextMatrix(X, 9)
   '                     .Fields(9) = Me.Grid.TextMatrix(X, 10)
   '                     .Fields(10) = Me.Grid.TextMatrix(X, 7)
   '                     .Fields(11) = Me.Grid.TextMatrix(X, 11)
   '                     .Update
   '                 End If
   '             Next
   '         End With
   '         SSTab1.Tab = 1
   '         txtcadastro(9).SetFocus
   '     Else
   '         Msgbox "Favor Preencer o campo!", vbInformation, "ZEUS"
   '     End If
   '     rsGCF.Close
   '     Set rsGCF = Nothing
   '
   '     rsGpf.Close
   '     Set rsGpf = Nothing
    'End If
    If CodLV <> 0 Then txtcadastro(0) = CodLV
    If CodLV <> 0 Then txtcadastro(8) = CodLV
    rsGrid.Close
    Set rsGrid = Nothing
    cnBanco.CommitTrans
    Msgbox "Dados gravados com sucesso", vbInformation, "Ok!"
    Exit Sub
TrataErro:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
        cnBanco.RollbackTrans
    End If
End Sub

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim y As Integer, x As Integer
    y = vListViewPrincipal.ListItems.Count
    For x = 1 To y
        If vListViewPrincipal.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = vListViewPrincipal.ListItems.Add(, , Format(txtcadastro(0), "000000"))
        ItemLst.SubItems(1) = txtcadastro(2).Text
        ItemLst.SubItems(2) = txtcadastro(3).Text
        ItemLst.SubItems(3) = txtcadastro(21).Text
        ItemLst.SubItems(4) = txtcadastro(4).Text
        ItemLst.SubItems(5) = txtcadastro(5).Text
        ItemLst.SubItems(6) = cboCadastro(0).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(7) = ""
            ItemLst.ListSubItems.Item(7).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(7) = ""
            ItemLst.ListSubItems.Item(7).ReportIcon = "OK"
        End If
    Else
        vListViewPrincipal.SelectedItem.ListSubItems.Item(1) = txtcadastro(2).Text
        vListViewPrincipal.SelectedItem.ListSubItems.Item(2) = txtcadastro(3).Text
        vListViewPrincipal.SelectedItem.ListSubItems.Item(3) = txtcadastro(21).Text
        vListViewPrincipal.SelectedItem.ListSubItems.Item(4) = txtcadastro(4).Text
        vListViewPrincipal.SelectedItem.ListSubItems.Item(5) = txtcadastro(5).Text
        vListViewPrincipal.SelectedItem.ListSubItems.Item(6) = cboCadastro(0).Text
        If Check1.Value = 0 Then
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7) = ""
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7).ReportIcon = "EXC"
        Else
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7) = ""
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    Msgbox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    'Para essa linha de comando existe um função dentro do módulo RotinaGeral
    'responsavel por desabilitar o BIP qdo precionada a tecla ENTER nos Texbox
    KeyAscii = Enter(KeyAscii)
    '-----------------
End Sub

Private Sub txtCadastro_LostFocus(Index As Integer)
    If Index = 8 Then
        Me.txtcadastro(1).SetFocus
    End If
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsCliFor As New ADODB.Recordset
    Dim sql As String
    sql = "Select top 1 * from tbclifor order by codclifor Desc"
    rsCliFor.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCliFor.RecordCount > 0 Then
        GeraCodigo = rsCliFor.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    rsCliFor.Close
    Set rsCliFor = Nothing
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

Private Sub ResultPesq()
    rsCliFor.MoveFirst
    rsCliFor.Find "codclifor=" & "'" & Val(varGlobal) & "'"
    If Not rsCliFor.EOF Then
        'If rsCliFor.Fields(10) = 1 Then
            rsCliForJ.MoveFirst
            rsCliForJ.Find "codclifor=" & "'" & Val(varGlobal) & "'"
        'Else
        '    rsCliForF.MoveFirst
        '    rsCliForF.Find "codclifor=" & "'" & Val(varGlobal) & "'"
        'End If
    End If
    CompoeControles
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    'If chkCadastro(0).Value = 0 And chkCadastro(1) = 0 Then
    '    Msgbox "Favor informar o campo Especificação!", vbInformation, "Atenção"
    '    Me.chkCadastro(0).SetFocus
    '    Exit Function
    'End If
    If txtcadastro(23).Text = "" Then
        Msgbox "Favor informar o campo Código do ramo de atividade!", vbInformation, "Atenção"
        Me.mskCadastro(10).SetFocus
        Exit Function
    End If
    If SSTab1.TabEnabled(0) = True Then
        If Me.txtcadastro(1) = "" Then
            Msgbox "Favor preencher o campo Razão Social!", vbInformation, "Atenção"
            Me.txtcadastro(1).SetFocus
            Exit Function
        ElseIf Me.txtcadastro(2) = "" Then
            Msgbox "Favor preencher o campo Nome Fantasia!", vbInformation, "Atenção"
            Me.txtcadastro(2).SetFocus
            Exit Function
        ElseIf Me.mskCadastro(2) = "" Then
            Msgbox "Favor preencher o campo Telefone", vbInformation, "Atenção"
            Me.mskCadastro(2).SetFocus
            Exit Function
        ElseIf Me.cboCadastro(0) = "" Then
            Msgbox "Favor preencher o campo Estado", vbInformation, "Atenção"
            Me.cboCadastro(0).SetFocus
            Exit Function
        End If
    ElseIf SSTab1.TabEnabled(0) = False Then
        If Me.txtcadastro(9) = "" Then
            Msgbox "Favor preencher o campo Nome", vbInformation, "Atenção"
            Me.txtcadastro(1).SetFocus
            Exit Function
        ElseIf Me.mskCadastro(5) = "" Then
            Msgbox "Favor preencher o campo Telefone", vbInformation, "Atenção"
            Me.mskCadastro(2).SetFocus
            Exit Function
        ElseIf Me.cboCadastro(1) = "" Then
            Msgbox "Favor preencher o campo Estado", vbInformation, "Atenção"
            Me.cboCadastro(1).SetFocus
            Exit Function
        End If
    End If
    ValidaCampo = True
End Function

Private Function ValidaItem(smensagem)
    Dim x As Byte
    If txtcadastro(17) = "" Then
        smensagem = "Favor Informar o nome do contato"
        Me.txtcadastro(17).SetFocus
        ValidaItem = False
        Exit Function
    End If
    
    If ByLinhaInclusaoGrid = 0 Then
        If Not VerificaGrid(txtcadastro(17).Text) Then
            smensagem = "Contato já digitado!"
            Me.txtcadastro(17).SetFocus
            ValidaItem = False
            Exit Function
        End If
    End If
    ValidaItem = True
    Exit Function
End Function

Private Function VerificaGrid(nomContato)
    Dim x As Byte

    For x = 1 To Me.Grid.Rows - 1
        If nomContato = Me.Grid.TextMatrix(x, 2) Then
            VerificaGrid = False
            Exit Function
        End If
    Next
    VerificaGrid = True
End Function

Private Function DesbloqueiaControles()
    Dim x As Integer
    
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = True
    Next
    For x = 0 To mskCadastro.Count - 1
        mskCadastro(x).Enabled = True
    Next
    For x = 0 To cboCadastro.Count - 1
        cboCadastro(x).Enabled = True
    Next
    For x = 0 To cmdCadastro.Count - 1
        cmdCadastro(x).Enabled = True
    Next
    txtcadastro(0).Enabled = False
    txtcadastro(8).Enabled = False
    Grid.Enabled = True
    Grid.ForeColor = &H80000008
End Function

Private Function BloqueiaControles()
    Dim x As Integer
    For x = 0 To txtcadastro.Count - 1
        txtcadastro(x).Enabled = False
    Next
    For x = 0 To mskCadastro.Count - 1
        mskCadastro(x).Enabled = False
    Next
    For x = 0 To cboCadastro.Count - 1
        cboCadastro(x).Enabled = False
    Next
    For x = 0 To cmdCadastro.Count - 1
        cmdCadastro(x).Enabled = False
    Next
    Grid.Enabled = False
    Grid.ForeColor = &H808080
End Function

Private Sub txtCadastro_GotFocus(Index As Integer)
    Dim x As Integer
    For x = 1 To txtcadastro.Count - 1
        txtcadastro(x).SelStart = 0
        txtcadastro(x).SelLength = Len(txtcadastro(x).Text)
    Next
End Sub

Private Sub Mskcadastro_GotFocus(Index As Integer)
    Dim x As Integer
    For x = 0 To mskCadastro.Count - 1
        mskCadastro(x).SelStart = 0
        mskCadastro(x).SelLength = Len(mskCadastro(x).Text)
    Next
End Sub


