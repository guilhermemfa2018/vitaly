VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8550
   ClientLeft      =   510
   ClientTop       =   1320
   ClientWidth     =   14235
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8175
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Data do sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5010
            MinWidth        =   5010
            Object.ToolTipText     =   "Usu�rio logado"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7505
            MinWidth        =   4304
            Object.ToolTipText     =   "Grupo do usu�rio logado"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6175
            MinWidth        =   6175
            Object.ToolTipText     =   "DB rede"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7920
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ZEUS.XTREMERibbon Ribbon 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
      Begin VB.Frame Frame7 
         Caption         =   "Par�metros do M�dulo Avaliador"
         Height          =   1695
         Left            =   2640
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkAvaliador 
            Caption         =   "Experi�ncia:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkAvaliador 
            Caption         =   "Habilidades:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkAvaliador 
            Caption         =   "Cursos/treinamentos:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox chkAvaliador 
            Caption         =   "Forma��o escolar:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Frame Frame10 
            Caption         =   "M�dia geral"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   735
            Left            =   2880
            TabIndex        =   6
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            Begin VB.Label Label41 
               Caption         =   "Label41"
               Height          =   255
               Left            =   360
               TabIndex        =   7
               Top             =   360
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   5160
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin MSMask.MaskEdBox mskCadMatriz 
            Height          =   285
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label37 
            Caption         =   "Label37"
            Height          =   255
            Left            =   2040
            TabIndex        =   17
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label38 
            Caption         =   "Label38"
            Height          =   255
            Left            =   2040
            TabIndex        =   16
            Top             =   840
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label39 
            Caption         =   "Label39"
            Height          =   255
            Left            =   2040
            TabIndex        =   15
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label40 
            Caption         =   "Label40"
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   1320
            Visible         =   0   'False
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1680
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   152
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3AFA
            Key             =   ""
            Object.Tag             =   "ramo de atividades"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":47D4
            Key             =   ""
            Object.Tag             =   "Clientes"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":54AE
            Key             =   ""
            Object.Tag             =   "transportadora"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6188
            Key             =   ""
            Object.Tag             =   "tipo de material"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6E62
            Key             =   ""
            Object.Tag             =   "materiais"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7B3C
            Key             =   ""
            Object.Tag             =   "itens de verifica��o"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8816
            Key             =   ""
            Object.Tag             =   "Projetos"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":94F0
            Key             =   ""
            Object.Tag             =   "processos"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A1CA
            Key             =   ""
            Object.Tag             =   "or�amentos"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AEA4
            Key             =   ""
            Object.Tag             =   "fce"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BB7E
            Key             =   ""
            Object.Tag             =   "lm"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C858
            Key             =   ""
            Object.Tag             =   "ld"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D532
            Key             =   ""
            Object.Tag             =   "os"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E20C
            Key             =   ""
            Object.Tag             =   "evolu��o"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EEE6
            Key             =   ""
            Object.Tag             =   "emitir relat�rio"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FBC0
            Key             =   ""
            Object.Tag             =   "Imprimir relat�rio"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1089A
            Key             =   ""
            Object.Tag             =   "configura��es"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11574
            Key             =   ""
            Object.Tag             =   "grupos"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1224E
            Key             =   ""
            Object.Tag             =   "usu�rios"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12F28
            Key             =   ""
            Object.Tag             =   "menu"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13C02
            Key             =   ""
            Object.Tag             =   "skin"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":148DC
            Key             =   ""
            Object.Tag             =   "fundo"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":155B6
            Key             =   ""
            Object.Tag             =   "Sistema"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16290
            Key             =   ""
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16F6A
            Key             =   ""
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17C44
            Key             =   ""
            Object.Tag             =   "Desenho"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1891E
            Key             =   ""
            Object.Tag             =   "Check"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":195F8
            Key             =   ""
            Object.Tag             =   "Controle"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1A2D2
            Key             =   ""
            Object.Tag             =   "pdf"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4E97C
            Key             =   ""
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4F656
            Key             =   ""
            Object.Tag             =   "Atualizar"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":50330
            Key             =   ""
            Object.Tag             =   "Cadastro"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5100A
            Key             =   ""
            Object.Tag             =   "Lista"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":51CE4
            Key             =   ""
            Object.Tag             =   "Baixar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":529BE
            Key             =   ""
            Object.Tag             =   "Baixar"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":53138
            Key             =   ""
            Object.Tag             =   "Cadastro"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":53E12
            Key             =   ""
            Object.Tag             =   "Cargos"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":54AEC
            Key             =   ""
            Object.Tag             =   "Configuracoes"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":557C6
            Key             =   ""
            Object.Tag             =   "Configuracoes"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":564A0
            Key             =   ""
            Object.Tag             =   "Dados"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5717A
            Key             =   ""
            Object.Tag             =   "Desenhos"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":57E54
            Key             =   ""
            Object.Tag             =   "fases"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":58B2E
            Key             =   ""
            Object.Tag             =   "escolar"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":59808
            Key             =   ""
            Object.Tag             =   "escolar"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5A4E2
            Key             =   ""
            Object.Tag             =   "escolar"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5B1BC
            Key             =   ""
            Object.Tag             =   "desenvolvimento"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5BE96
            Key             =   ""
            Object.Tag             =   "Orcamento"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5CB70
            Key             =   ""
            Object.Tag             =   "programacao"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5D84A
            Key             =   ""
            Object.Tag             =   "programacao"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5E524
            Key             =   ""
            Object.Tag             =   "treinamento"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5F1FE
            Key             =   ""
            Object.Tag             =   "Zeus"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5FED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":60BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6188C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":62566
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":63240
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":63F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":64BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":658CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":665A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":67282
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":67F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":68C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":69910
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6A5EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6B2C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6BF9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6CC78
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6D952
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6E62C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6F306
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6FFE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":70CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7A89A
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7B574
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7C24E
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7CF28
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7DC02
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7E8DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7F5B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":80290
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":80F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":81C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8291E
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":835F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":842D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":84FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":85C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":86960
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8763A
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":88314
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":88FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":89CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8A9A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8B67C
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8C356
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8D030
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8DD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8E9E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8F6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":90398
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":91072
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":91D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":92A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":93700
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":943DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":950B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":95D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":96A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":97742
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9841C
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":990F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9ADD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A873F
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A9419
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AA0F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AADCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":ABAA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AC781
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AD45B
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AE135
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AEE0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AFAE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B07C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B149D
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B2177
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B2E51
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B3B2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B4805
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B54DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B61B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B6E93
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B7B6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B8847
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B9521
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BA1FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BAED5
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BBBAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BC889
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BD563
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BE23D
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BEF17
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BFBF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C08CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C15A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C227F
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C2F59
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C3C33
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C490D
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C55E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C62C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C6F9B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14175
      TabIndex        =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   14235
      Begin VB.Image Image1 
         Height          =   11520
         Left            =   2280
         Picture         =   "Principal.frx":C7C75
         Top             =   0
         Width           =   20400
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "Principal.frx":CFF07
      Top             =   6360
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   2280
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   152
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D013B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E04C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":ECDF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F9724
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":106056
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":115B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":124A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1313A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13DCD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14FBCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15F5C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16BEF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":178829
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":18515B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":191A8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":19E3BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1AACF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1B7623
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1C3F55
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1D0887
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1DD1B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1E9AEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1F641D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":202D4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":212264
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":21EB96
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":22B4C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":23A899
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2471CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":253AFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":26042F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":26CD61
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":27CE9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2897CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2977DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2A410C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2B4087
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2C6CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2D35ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2DFF1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2EC851
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":2FC303
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":308C35
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":315567
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":321E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":32E7CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":33B0FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":349E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":35677A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3630AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":36F9DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":38389C
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3901CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":39CB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3A9432
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3B5D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3C2696
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3CEFC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3DB8FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3E822C
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3FB534
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":407E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":414798
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4210CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":42D9FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":43A32E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":446C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":453592
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":45FEC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":46C7F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":479128
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":485A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":49238C
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":49ECBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4AB5F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4B7F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4C86A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4D76C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4E3FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4F0926
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4FD258
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":50BF22
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":518854
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5298AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5361E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":542B13
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":54F445
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":55BD77
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5686A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":574FDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":58190D
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":58E23F
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":59AB71
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5A74A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5B3DD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5C0707
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5D201A
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5DE94C
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5EB27E
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":5F7BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6044E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":610E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":621F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":62E876
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":63B1A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":647ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":65440C
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":660D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":66D670
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":679FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6868D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":693206
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":69FB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6AC46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6BE757
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6CB089
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6D79BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6E42ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6F0C1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":6FD551
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":709E83
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7167B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7230E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":72FA19
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":73C34B
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":74DBB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":75A4E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":766E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":77374A
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":786AE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":793419
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":79FD4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7AF779
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7BC0AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7D07FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7E317B
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7EFAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":7FC3DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":808D11
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":81789C
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8241CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":830B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":83D432
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":84E08A
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":85A9BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":86A983
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8772B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":887703
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8969E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8A6178
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8B69E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8C65F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   2880
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   152
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8D6203
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8E2B35
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8EF467
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8FBD99
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9086CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":917826
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":924158
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":930A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":93D3BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":949CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":956620
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":962F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":96F884
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":97C1B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":988AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":99541A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9A1D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9AE67E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9BAFB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9C78E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9D4214
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9E0B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9ED478
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9F9DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A066DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A1300E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A1F940
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A2C272
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A38BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A454D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A51E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A5E73A
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A6B06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A7799E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A842D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A90C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A9D534
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AA9E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AB6798
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AC30CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":ACF9FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":ADC32E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AE8C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":AF5592
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B01EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B0E7F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B1B128
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B27A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B3438C
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B40CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B4D5F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B59F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B66854
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B73186
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B7FAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B8C3EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B98D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BA564E
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BB1F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BBE8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BCB1E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BDC429
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BE8D5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":BF568D
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C01FBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C0E8F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C1B223
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C2B038
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C3796A
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C4429C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C50BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C5D500
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C69E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C76764
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C83096
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C91962
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CA05B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CACEE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CB9817
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CC6149
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CD2A7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CE2585
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CEEEB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CFB7E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D0811B
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D14A4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D2137F
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D2DCB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D3A5E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D46F15
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D53847
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D60179
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D6CAAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D793DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D85D0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":D92641
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DA310B
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DAFA3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DBC36F
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DC8CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DD55D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DE1F05
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DEE837
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DFB169
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E07A9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E143CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E20CFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E2D631
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E39F63
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E46895
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E531C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E5FAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E6C42B
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E78D5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E8568F
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E91FC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EA0C6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EAD59D
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EB9ECF
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EC6801
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":ED3133
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EDFA65
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EEC397
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EF8CC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F055FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F11F2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F1E85F
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F2B191
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F37AC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F443F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F50D27
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F5D659
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F69F8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F768BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F8941E
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F95D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FA2682
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FAEFB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FBB8E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FC8218
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FD4B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FE147C
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FEDDAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FFA6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1007012
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1013944
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1020276
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":102CBA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10394DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1045E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":105273E
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":105F070
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList6 
      Left            =   3480
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   152
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":106B9A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10782D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1084C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1091538
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":109DE6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10AA79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10B70CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10C3A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10D0332
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10DCC64
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10E9596
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10F5EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11027FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":110F12C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":111BA5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1128390
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1134CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11415F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":114DF26
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":115A858
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":116718A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1173ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11803EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":118CD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1199652
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11A5F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11B28B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11BF1E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11CBB1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11D844C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11E4D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11F16B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11FDFE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":120A914
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1217246
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1223B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12304AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":123CDDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":124970E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1256040
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1262972
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":126F2A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":127BBD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1288508
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1294E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12A176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12AE09E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12BA9D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12C7302
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12D3C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12E0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12ECE98
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12F97CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13060FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1312A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":131F360
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":132BC92
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13385C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1344EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1351828
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":135E15A
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":136AA8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13773BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1383CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1390622
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":139CF54
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13A9886
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13B969B
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13C5FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13D28FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13DF231
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13EBB63
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13F8495
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1404DC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14116F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":141E02B
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":142A95D
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":143728F
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1443BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14504F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":145CE25
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1469757
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1476089
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14829BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":148F2ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":149BC1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14A8551
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14B4E83
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14C17B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14CE0E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14DAA19
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14E734B
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14F3C7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15005AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":150CEE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1519813
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1526145
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1532A77
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":153F3A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":154BCDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":155860D
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1564F3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1571871
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":157E1A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":158AAD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1597407
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15A3D39
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15B066B
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15BCF9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15C98CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15D6201
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15E2B33
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15EF465
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15FBD97
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16086C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1614FFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":162192D
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":162E25F
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":163AB91
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16474C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1653DF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1660727
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":166D059
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":167998B
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16862BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1692BEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":169F521
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16ABE53
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16B8785
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16C50B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16D19E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16DE31B
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16EAC4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":16F757F
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1703EB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17107E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":171D115
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1729A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1736379
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1742CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":174F5DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":175BF0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1768841
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1775173
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1781AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":178E3D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":179AD09
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17A763B
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17B3F6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17C089F
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17CD1D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":17D9B03
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tema As String
Private rsConf As New ADODB.Recordset
Private SqlConf As String
Private vFechar As Integer
Private rsCandidatos As New ADODB.Recordset
Private sqlCandidatos As String

Sub EstendeImagem()
    picBackdrop.Cls
    picBackdrop.Visible = True
    picBackdrop.AutoRedraw = True
    picBackdrop.BackColor = &H8000000C
    picBackdrop.Height = Me.Height
    Image1.Stretch = True
    Image1.Top = 0
    Image1.Left = 0
    Image1.Height = picBackdrop.Height
    Image1.Width = picBackdrop.Width
    picBackdrop.PaintPicture Image1, Image1.Left, Image1.Top, Image1.Width, Image1.Height
    Principal.Picture = picBackdrop.Image
    picBackdrop.Visible = False
End Sub

Private Function AlteraRibon()
Tema = Tema + (1)
If Tema = 19 Then Tema = 0
Ribbon.Theme = Tema
Ribbon.Refresh

'Salva o Tema atual
WriteProfile "Tema", "NomeTema", Tema, App.Path & "\CONFIG.INI"

End Function

Private Sub MDIForm_Activate()
    MDIForm_Resize
End Sub

'Faz a imagem caber no formul�rio MDI
Private Sub MDIForm_Resize()
    On Error Resume Next
    Set Image1.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")
    EstendeImagem
    If Principal.WindowState = 2 And vStatusWin = 2 Then Principal.WindowState = 2
    If Principal.WindowState = 0 And vStatusWin = 2 Then
        Principal.WindowState = 1
        vStatusWin = 1
    End If
    If Principal.WindowState = 0 And vStatusWin = 1 Then
        Principal.WindowState = 2
        vStatusWin = 2
    End If
    
End Sub

Private Sub MDIForm_Load()

    'Exit Sub
    
    'On Error GoTo ErrHandler
    'Recupera o Tema atual
    vStatusWin = 2
    LimiteLinhas = 500 ' Val(Text1.Text)
    
    Tema = GetValue(App.Path & "\CONFIG.ini", "Tema", "NomeTema", "")
    
    'Pega o Skin atual salvo na pasta principal com o nome MySkin
    Skin1.LoadSkin App.Path & "\MySkin.skn"
    Skin1.ApplySkin Me.HWnd
    
    Me.Caption = "ZEUS - Sistema de Controle de Produ��o" & " - Vers�o: " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Pega a imagem de funco atual salva na pasta principal com o nome PlanoDeFundo
    Set Principal.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")

    '### >> Aqui come�a o Ribbon << ############################################################
    '# SET Theme BEFORE ALL
    Ribbon.Theme = Tema
    
    '# Set ImageList to use for icons
    If vColectionIcons = 1 Then
        Ribbon.ImageList = ImageList3
    ElseIf vColectionIcons = 2 Then
        Ribbon.ImageList = ImageList4
    ElseIf vColectionIcons = 3 Then
        Ribbon.ImageList = ImageList5
    ElseIf vColectionIcons = 4 Then
        Ribbon.ImageList = ImageList6
    ElseIf vColectionIcons = 5 Then
        Ribbon.ImageList = ImageList3
    End If
    '# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
    Ribbon.ButtonCenter = False

    abreConfMenu
    montaMenu
    fechaConfMenu
    montaTabMenu
    Ribbon.Refresh

    StatusBar1.Panels(1).Width = 1840
    StatusBar1.Panels(2).Width = 4440.189
    StatusBar1.Panels(1).Text = Format(Date, "dd/mm/yyyy")
    StatusBar1.Panels(2).Text = "Usu�rio: " & NomUsu
    StatusBar1.Panels(3).Text = "Grupo: " & GrupoUsu
    StatusBar1.Panels(4).Text = "DB: " & sServerName & " (" & sDatabaseName & ")"
Exit Sub
ErrHandler:
    Msgbox "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte T�cnico.", vbCritical, "Aten��o"
End Sub

Private Sub Ribbon_ButtonClick(ByVal ID As String, ByVal Caption As String)
    'On Error Resume Next
    Pesquisa = ""
    vControlaDim = 0
    checaFiltro = True
    'vAcaoTab = "OPEN"
    
    If ID = 1 Then  '(Movimenta��es OS - Paradas)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Paradas - OS"
        apontaLV = 2
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 2
        FiltroGeral = "Ativos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    If ID = 2 Then '(Clientes)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Clientes"
        apontaLV = 1
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 1
        FiltroGeral = "Ativos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 3 Then '(Transportadoras)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Transportadoras"
        apontaLV = 3
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 3
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 4 Then '(Tipo de Material)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Tipo de Material"
        apontaLV = 0
        If constroiTabs(frmPesqGeralTeste2.SSTab1, True) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 0
        'vListViewPrincipal.CheckBoxes = True
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    If ID = 5 Then '(Materiais)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "F�rmula de Produtos"
        apontaLV = 4
        If constroiTabs(frmPesqGeralTeste2.SSTab1, True) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 4
        'vListViewPrincipal.CheckBoxes = True
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    
    If ID = 6 Then '(Item Verifica��o)
        Set chamaForm = New frmConfSistema
        frmItemVerif.Show 1
    End If
    If ID = 7 Then '(Projetos)
        Set chamaForm = New frmProjetos
        frmProjetos.Show 1
    End If

    If ID = 8 Then '(Processos)
        Set chamaForm = New frmProcessos
        frmProcessos.Show 1
    End If
    
    If ID = 9 Then '(Desenhos)
        vQdtFrom = 3 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Desenhos"
        apontaLV = 7
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 7
        FiltroGeral = "Ativos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 10 Then '(F�rmula - Centro de Custo)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "F�rmula - Centro de Custo"
        apontaLV = 11
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 11
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 11 Then '(FO - Ficha de Or�amento - CADASTRO INICIAL DA FCE)
        vQdtFrom = 4 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "FO"
        apontaLV = 5
        If constroiTabs(frmPesqGeralTeste2.SSTab1, True) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 5
        'vListViewPrincipal.CheckBoxes = True
        FiltroGeral = "Ativos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    If ID = 12 Then '(Faturamento por FCE)
        vQdtFrom = 7 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Faturamento por FCE"
        apontaLV = 20
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 20
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 21 Then '(FCE - Ficha de Controle de Encomenda)
        vQdtFrom = 3 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "FCE"
        apontaLV = 6
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 6
        FiltroGeral = "Ativos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    If ID = 22 Then '(LM - Lista de Materiais)
        vQdtFrom = 3 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "LM"
        apontaLV = 8
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 8
        FiltroGeral = "Ativos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    If ID = 23 Then '(MP - M�todos e Processos)
        vQdtFrom = 3 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "MP"
        apontaLV = 9
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 9
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    
    If ID = 26 Then '(CD - Controle de Desenhos)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Controle de Desenhos"
        apontaLV = 10
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 10
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 27 Then '(Monitorar Produ��o)
        frmMonitorar.Show
    End If

    If ID = 31 Then ' Qualidade (RNCF - Registro de N�o Conformidade de Fabrica��o)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "RNCF"
        apontaLV = 12
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 12
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    If ID = 32 Then
        frmComunicacaoDesvio.Show 1
    End If
    
    If ID = 35 Then 'Relat�rios de Inspe��o
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Relat�rio de Inspe��o"
        apontaLV = 16
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 16
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If
    
    If ID = 36 Then 'Impress�o de Relat�rios de Inspe��o
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Imp. Rel. de Inspe��o"
        apontaLV = 19
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 19
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 41 Then ' Emiss�o de Relat�rios de Expedi��o
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Relat�rio de Expedi��o"
        apontaLV = 17
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 17
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 42 Then ' Impress�o de Relat�rios de Expedi��o
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Imp. Rel. de Expedi��o"
        apontaLV = 18
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 18
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 51 Then '(Sistema)
        'Unload frmPesqGeralTeste2
        Set chamaForm = New frmConfSistema
        frmConfSistema.Show 1
    End If

    If ID = 52 Then '(Grupos)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Grupos"
        apontaLV = 14
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 14
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
        'tabAberta = True
    End If
    If ID = 53 Then '(Usu�rios)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Usu�rios"
        apontaLV = 13
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 13
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

'---------- Configura��es de ambiente
    If ID = 54 Then
        AlteraRibon
    End If
    If ID = 55 Then
        FrmSkins.Show
        Exit Sub
    End If
    If ID = 56 Then
        frmLocalizar.Show vbModal
    End If
'----------
    If ID = 57 Then
        vQdtFrom = 3 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "OS Permiss�es"
        apontaLV = 15
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 15
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 58 Then '(Terceirizados)
        vQdtFrom = 1 ''Especifica a quantidade de FROM na query do filtro
        Formulario = "Terceiros"
        apontaLV = 21
        If constroiTabs(frmPesqGeralTeste2.SSTab1, False) = True Then
            'contruirBotoesPorModulo apontaLV
            Exit Sub
        End If
        DimensionaLV1 "M�todos e Processos", vFramePrincipal, vListViewPrincipal, vLabelPrincipal
        apontaLV = 21
        FiltroGeral = "Todos"
        frmPesqGeralTeste2.Timer1.Enabled = True
        MudaPropPictureTeste frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal
    End If

    If ID = 61 Then '(Sobre)
        frmRegistro.Show 1
    End If

    If ID = 62 Then '(Ajuda)
        LoadEXE (App.Path & "\ZEUSHHelp.exe")
    End If

    If ID = 71 Then '(Reabertura de OS)
        frmReabrirOP.Show
    End If
    
'    If MudaPropPictureTeste(frmPesqGeralTeste2.SSTab1.Tab, vListViewPrincipal) = False Then 'Configura Picture para colorir as linhas do listview de acordo com o Tipo de FCE
'        Unload frmPesqGeralTeste2
'        Set frmPesqGeralTeste2 = Nothing
'    End If

'''    If ID = 1 Then  '(Movimenta��es OS - Paradas)
'''        apontaLV = 2
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 2 Then '(Clientes)
'''        apontaLV = 1
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 3 Then '(Transportadora)
'''        apontaLV = 3
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 4 Then '(Tipo de Material)
'''        apontaLV = 0
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 5 Then '(Materiais)
'''        MeuLV.ListView1.CheckBoxes = True
'''        apontaLV = 4
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 8 Then '(Processos)
'''        Set chamaForm = New frmProcessos
'''        frmProcessos.Show 1
'''    End If
'''    If ID = 9 Then '(Desenhos)
'''        apontaLV = 7
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 10 Then '(F�rmula - Centro de Custo)
'''        apontaLV = 11
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 11 Then '(FO - Ficha de Or�amento - CADASTRO INICIAL DA FCE)
'''        MeuLV.ListView1.CheckBoxes = True
'''        apontaLV = 5
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 12 Then '(Faturamento por FCE)
'''        apontaLV = 20
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 21 Then '(FCE - Ficha de Controle de Encomenda)
'''        apontaLV = 6
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 26 Then '(CD - Controle de Desenhos)
'''        apontaLV = 10
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 27 Then '(Monitorar Produ��o)
'''        frmMonitorar.Show
'''    End If
'''
''''    If ID = 21 Then '(Programa��o)
''''        MeuLV.ListView1.CheckBoxes = True
''''        FiltroGeral = "Ativos pendentes"
''''        apontaLV = 10
''''        MontaLV (apontaLV)
''''        'MeuLV.ListView1.Checkboxes = False
''''    End If
'''    If ID = 22 Then '(LM - Lista de Materiais)
'''        apontaLV = 8
'''        FiltroGeral = "Ativos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 23 Then '(MP - M�todos e Processos)
'''        apontaLV = 9
'''        FiltroGeral = "Todos"
'''        AplicarSkin Me, Principal.Skin1
'''        NewColorDBGrid Me
'''
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 31 Then ' Qualidade (RNCF - Registro de N�o Conformidade de Fabrica��o)
'''        apontaLV = 12
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 32 Then
'''        frmComunicacaoDesvio.Show 1
'''    End If
'''    If ID = 33 Then
'''        'FCRTreinCargo.Show 1
'''    End If
'''
'''    If ID = 35 Then 'Relat�rios de Inspe��o
'''        apontaLV = 16
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 36 Then 'Impress�o de Relat�rios de Inspe��o
'''        apontaLV = 19
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 41 Then ' Emiss�o de Relat�rios de Expedi��o
'''        apontaLV = 17
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 42 Then ' Impress�o de Relat�rios de Expedi��o
'''        apontaLV = 18
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 51 Then '(Sistema)
'''        'Principal.aicAlphaImage1.Visible = True
'''        Set chamaForm = New frmConfSistema
'''        frmConfSistema.Show 1
'''        'Principal.aicAlphaImage1.Visible = False
'''    End If
'''    If ID = 52 Then '(Grupos)
'''        apontaLV = 14
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''    If ID = 53 Then '(Usu�rios)
'''        apontaLV = 13
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
''''---------- Configura��es de ambiente
'''    If ID = 54 Then
'''        AlteraRibon
'''    End If
'''    If ID = 55 Then
'''        FrmSkins.Show
'''        Exit Sub
'''    End If
'''    If ID = 56 Then
'''        frmLocalizar.Show vbModal
'''    End If
''''----------
'''    If ID = 57 Then
'''        apontaLV = 15
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''    If ID = 58 Then '(Terceirizados)
'''        apontaLV = 21
'''        FiltroGeral = "Todos"
'''        MontaLV (apontaLV)
'''    End If
'''
'''
'''    If ID = 71 Then '(Reabertura de OS)
'''        frmReabrirOP.Show
'''        'frmRegistro.Show 1
'''    End If
'''
'''
'''    If ID = 61 Then '(Sobre)
'''        frmRegistro.Show 1
'''    End If
'''
'''    If ID = 62 Then '(Ajuda)
'''        LoadEXE (App.Path & "\ZEUSHHelp.exe")
'''    End If
End Sub

Private Function FecharPrograma()
End
End Function

Private Sub mnuSair_Click()
End
End Sub
   
Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuConsig_Click()
'Consignacao.Show
End Sub
Private Sub mnuNvComp_Click()
'AgendarCompromissos.Show
End Sub
Private Sub mnuCodeBar_Click()
'CODEBAR.Show
End Sub
Private Sub mnuCalc_Click()
'AbreCalculadora
End Sub
Private Sub mnuAjuda_Click()
'    mobjMsg.Abrir "Ajuda em constru��o, aguarde.", , informacao, "Master System"
End Sub
Private Sub mnuFechaJanelas_Click()
    FechaJanelas
End Sub
Sub FechaJanelas()
    Dim Frm As Form
    For Each Frm In Forms
        If Frm.Name <> Me.Name Then
           'fecha todas as telas exceto a chamadora (MDI Form)
           Unload Frm
        End If
    Next Frm
End Sub
Private Sub mnuLDLHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuLDLVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuOrganizaIcones_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub LoadEXE(Dir As String)
On Error GoTo ERRO
    Dim X As Integer
    Dim nofreeze As Integer
    X = Shell(Dir, 1)
    nofreeze = DoEvents()
    Exit Sub
ERRO:
    If Err.Number = 6 Then Exit Sub
   Msgbox "Arquivo de HELP n�o foi localizado !!! Verifique sua localiza��o ...", vbCritical, "Aten��o"
End Sub

Private Sub montaTabMenu()
On Error GoTo Err
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsCopia As New ADODB.Recordset
    Dim sqlCopia As String
    
    
    Dim rsMenuExpert As New ADODB.Recordset
    Dim sqlMenuExpert As String
    
    sqlMenuExpert = "Select * from tbMenuConf order by idsub"
    rsMenuExpert.Open sqlMenuExpert, cnBanco, adOpenKeyset, adLockReadOnly
10  cnBanco.BeginTrans
    sqlDeletar = "Delete from tbMenu"
    rsDeletar.Open sqlDeletar, cnBanco
    
    If rsMenuExpert.RecordCount > 0 Then
        sqlCopia = "Select * into tbConfGrupoCOPIA from tbConfGrupo"
        rsCopia.Open sqlCopia, cnBanco

        sqlDeletar = "Delete from tbConfGrupo where tbconfgrupo.tipo <> 'CHK' and tbconfgrupo.idgrupo = '" & XCodGrp & "'"
        rsDeletar.Open sqlDeletar, cnBanco
        While Not rsMenuExpert.EOF
            SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values('" & rsMenuExpert.Fields(0) & "','" & rsMenuExpert.Fields(1) & "','" & rsMenuExpert.Fields(2) & "','" & rsMenuExpert.Fields(3) & "','" & rsMenuExpert.Fields(5) & "')"
            rsMenu.Open SqlMenu, cnBanco

            SqlMenu = "Insert into tbConfGrupo(idgrupo,idmenu,idsub,tipo,nome,status,codcoligada,icon) Values(" & XCodGrp & ",'" & rsMenuExpert.Fields(0) & "','" & rsMenuExpert.Fields(1) & "','" & rsMenuExpert.Fields(2) & "','" & rsMenuExpert.Fields(3) & "','S','" & rsMenuExpert.Fields(5) & "','" & rsMenuExpert.Fields(6) & "')"
            rsMenu.Open SqlMenu, cnBanco

            rsMenuExpert.MoveNext
        Wend
        rsMenuExpert.Close
        Set rsMenuExpert = Nothing

        'Restaurando Permiss�es
        sqlCopia = "Select * from tbConfGrupoCOPIA"
        rsCopia.Open sqlCopia, cnBanco, adOpenKeyset, adLockReadOnly
        While Not rsCopia.EOF
            SqlMenu = "Update tbConfGrupo set status = '" & rsCopia.Fields(5) & "',incluir = '" & rsCopia.Fields(9) & "',editar = '" & rsCopia.Fields(10) & "',excluir = '" & rsCopia.Fields(11) & "',salvar = '" & rsCopia.Fields(12) & "',imprimir = '" & rsCopia.Fields(13) & "',filtrar = '" & rsCopia.Fields(14) & "' where idgrupo = '" & rsCopia.Fields(0) & "' and idmenu = '" & rsCopia.Fields(1) & "' and idsub = '" & rsCopia.Fields(2) & "'"
            rsMenu.Open SqlMenu, cnBanco
            rsCopia.MoveNext
        Wend
        rsCopia.Close
        Set rsCopia = Nothing

        sqlCopia = "Drop table tbConfGrupoCOPIA"
        rsCopia.Open sqlCopia, cnBanco
    
    Else
        SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','TAB','Cadastros','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','CAT','Prim�rios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'02','CAT','Secund�rios','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0101','BUT','Ramo de atividades','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0102','BUT','Clientes','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0103','BUT','Transportadoras','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0104','BUT','Tipo material','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0205','BUT','Materiais','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0206','BUT','Itens verifica��o','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0207','BUT','Projetos','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0208','BUT','Processos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'02','TAB','Or�amentos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'11','CAT','Vendas','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'1111','BUT','Servi�os','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'03','TAB','Planejamento','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'21','CAT','Planejamento e Controle da Produ��o','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2121','BUT','FCE','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2122','BUT','LM','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2123','BUT','LD','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2124','BUT','OS','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2125','BUT','Controle de Desenhos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'04','TAB','Produ��o','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'31','CAT','Acompanhamento de Produ��o','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'3131','BUT','OS Acompanhamento','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'3132','BUT','Evolu��o','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'05','TAB','Inspe��o/Expedi��o','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'41','CAT','Emiss�o de Relat�rios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'4141','BUT','Emitir relat�rio','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'4142','BUT','Imprimir relat�rio','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'06','TAB','Configura��es','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'51','CAT','Parametriza��es','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'52','CAT','Apar�ncia','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5151','BUT','Sistema','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5152','BUT','Grupos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5153','BUT','Usu�rios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5254','BUT','Menu','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5255','BUT','Skin','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5256','BUT','Fundo','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'07','TAB','Sobre','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'61','CAT','Sobre','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'6161','BUT','Sobre ZEUS','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,'6162','BUT','Ajuda do ZEUS','" & vCodcoligada & "');"
        
        rsMenu.Open SqlMenu, cnBanco
    End If
    cnBanco.CommitTrans
    Set rsMenu = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub abreConfMenu()
On Error GoTo Err
'    SqlConf = "Select * from tbconfgrupo Where tbconfgrupo.codcoligada = '" & vCodcoligada & "' and tbconfgrupo.idgrupo = '" & XCodGrp & "'order by id"
    SqlConf = "Select * from tbconfgrupo Where tbconfgrupo.idgrupo = '" & XCodGrp & "' and codcoligada = " & vCodcoligada & " order by idsub"
    rsConf.Open SqlConf, cnBanco, adOpenKeyset, adLockReadOnly
    Exit Sub
Err:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub fechaConfMenu()
    rsConf.Close
    Set rsConf = Nothing
End Sub

Private Sub montaMenu()
    Dim vMenu As String
    While Not rsConf.EOF
        If rsConf.Fields(5) = "S" Then
            If rsConf.Fields(3) <> "CHK" Then
                If rsConf.Fields(3) = "TAB" Then
                    Ribbon.AddTab rsConf.Fields(1), rsConf.Fields(4)
                End If
                If rsConf.Fields(3) = "CAT" Then
                    Ribbon.AddCat Right$(rsConf.Fields(2), 2), rsConf.Fields(1), rsConf.Fields(4), False
                End If
                If rsConf.Fields(3) = "BUT" Then
                    If Len(rsConf.Fields(2)) = 4 Then
                        Ribbon.AddButton Right$(rsConf.Fields(2), 2), Mid$(rsConf.Fields(2), 1, 2), rsConf.Fields(4), rsConf.Fields(8)
                    Else
                        vMenu = Val(Mid$(rsConf.Fields(2), 3, 3))
                        If Len(vMenu) <> 3 Then
                            Ribbon.AddButton Right$(rsConf.Fields(2), 2), Mid$(rsConf.Fields(2), 4, 2), rsConf.Fields(4), rsConf.Fields(8)
                        Else
                            Ribbon.AddButton Right$(rsConf.Fields(2), 3), Mid$(rsConf.Fields(2), 3, 3), rsConf.Fields(4), rsConf.Fields(8)
                        End If
                    End If
                End If
            End If
        End If
        rsConf.MoveNext
    Wend
    Ribbon.Refresh
End Sub

Private Function atualizaCandidatos()
On Error GoTo Err
    'FILTRA
    '1 = Colaborador
    '2 = Candidato
    atualizaCandidatos = True
    Dim rsDeletaTemp As New ADODB.Recordset
    Dim sqlDeletaTemp As String
    Dim rsCandidatos As New ADODB.Recordset
    Dim sqlCandidatos As String
    
    sqlDeletaTemp = "delete from ##Tempglobal"
    rsDeletaTemp.Open sqlDeletaTemp, cnBanco

    sqlCandidatos = "select a.id,a.cpf,a.nomecolaborador,d.nomedepartamento,e.nomesetor,c.codmatriz,f.nomecargo from tbcolaboradores as a inner join tbcolaboradoreshist as b " & _
    "on a.codcoligada = '" & vCodcoligada & "' and a.ativo = 'S' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join " & _
    "tbdepartamentos as d on c.coddepartamento = d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor " & _
    "inner join tbcargos as f on c.codcargo = f.codcargo order by a.id"
    rsCandidatos.Open sqlCandidatos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCandidatos.RecordCount = 0 Then
        rsCandidatos.Close
        Set rsCandidatos = Nothing
        atualizaCandidatos = False
        Exit Function
    End If
    
    If Not rsCandidatos.EOF Then
        While Not rsCandidatos.EOF '.Move(Val(Combo1.Text))
            txtCadMatriz(4) = rsCandidatos.Fields(5) ' Matriz
            Text1 = rsCandidatos.Fields(5) & rsCandidatos.Fields(6) ' Matrix+nome do cargo
            chkAvaliador(0).Value = 0
            chkAvaliador(1).Value = 0
            chkAvaliador(2).Value = 0
            chkAvaliador(3).Value = 0
            'For X = 0 To Len(rsCandidatos.Fields(5))
                chkAvaliador(0).Value = 1
                chkAvaliador(1).Value = 1
                chkAvaliador(2).Value = 1
                chkAvaliador(3).Value = 1
            'Next
            mskCadMatriz = rsCandidatos.Fields(1) ' CPF
            Avaliador "colaborador"
            GravaColaboradores
            rsCandidatos.MoveNext
        Wend
    End If
    rsCandidatos.Close
    Set rsCandidatos = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Function

Private Sub GravaColaboradores()
On Error GoTo Err
    Dim rsGravaColaboradores As New ADODB.Recordset
    Dim sqlGravaColaboradores As String
    Dim vIdent As Integer
    vIdent = rsCandidatos.Fields(0)
    
    sqlGravaColaboradores = "INSERT INTO ##Tempglobal(id,cpf,nomecolaborador,departamento,setor,experiencia,habilidade,treinamento,formacao) VALUES('" & rsCandidatos.Fields(0) & "','" & rsCandidatos.Fields(1) & "','" & rsCandidatos.Fields(2) & "','" & rsCandidatos.Fields(3) & "','" & rsCandidatos.Fields(4) & "','" & Replace(RemoveMask(Label37), ",", ".") & "','" & Replace(RemoveMask(Label38), ",", ".") & "','" & Replace(RemoveMask(Label39), ",", ".") & "','" & Replace(RemoveMask(Label41), ",", ".") & "')"
    rsGravaColaboradores.Open sqlGravaColaboradores, cnBanco
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

Private Sub criaTabTemp()
On Error GoTo Err
    'Criando uma tabela tempor�ria global
    Dim rsTabTemp As New ADODB.Recordset
    Dim SqlTabTemp As String
    SqlTabTemp = "CREATE TABLE ##Tempglobal(id INT NOT NULL,CPF VARCHAR(50) NOT NULL,nomecolaborador VARCHAR(100) NOT NULL,departamento VARCHAR(100) NOT NULL, setor VARCHAR(100) NOT NULL, experiencia FLOAT NOT NULL, habilidade FLOAT NOT NULL, treinamento FLOAT NOT NULL, formacao FLOAT NOT NULL)"
    rsTabTemp.Open SqlTabTemp, cnBanco
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




