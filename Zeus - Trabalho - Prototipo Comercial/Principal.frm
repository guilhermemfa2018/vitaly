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
            Object.ToolTipText     =   "Usuário logado"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7505
            MinWidth        =   4304
            Object.ToolTipText     =   "Grupo do usuário logado"
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
         Caption         =   "Parâmetros do Módulo Avaliador"
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
            Caption         =   "Experiência:"
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
            Caption         =   "Formação escolar:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Frame Frame10 
            Caption         =   "Média geral"
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
            Object.Tag             =   "itens de verificação"
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
            Object.Tag             =   "orçamentos"
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
            Object.Tag             =   "evolução"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":EEE6
            Key             =   ""
            Object.Tag             =   "emitir relatório"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":FBC0
            Key             =   ""
            Object.Tag             =   "Imprimir relatório"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1089A
            Key             =   ""
            Object.Tag             =   "configurações"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11574
            Key             =   ""
            Object.Tag             =   "grupos"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1224E
            Key             =   ""
            Object.Tag             =   "usuários"
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

'Faz a imagem caber no formulário MDI
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

Me.Caption = "ZEUS - Sistema de Controle de Produção" & " - Versão: " & App.Major & "." & App.Minor & "." & App.Revision

'Pega a imagem de funco atual salva na pasta principal com o nome PlanoDeFundo
Set Principal.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")

'### >> Aqui começa o Ribbon << ############################################################
'# SET Theme BEFORE ALL
Ribbon.Theme = Tema

'# Set ImageList to use for icons
Ribbon.ImageList = ImageList3

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
    StatusBar1.Panels(2).Text = "Usuário: " & NomUsu
    StatusBar1.Panels(3).Text = "Grupo: " & GrupoUsu
    StatusBar1.Panels(4).Text = "DB: " & sServerName & " (" & sDatabaseName & ")"
Exit Sub
ErrHandler:
    Msgbox "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", vbCritical, "Atenção"
End Sub

Private Sub Ribbon_ButtonClick(ByVal ID As String, ByVal Caption As String)
    'On Error Resume Next
    Pesquisa = ""
    'MeuLV.cmdconsulta(9).Visible = False
    vControlaDim = 0
    Tipo = True
    checaFiltro = True
    
    If ID = 1 Then  '(Movimentações OS - Paradas)
        apontaLV = 2
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 2 Then '(Clientes)
        apontaLV = 1
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 3 Then '(Transportadora)
        apontaLV = 3
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If
    If ID = 4 Then '(Tipo de Material)
        apontaLV = 0
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 5 Then '(Materiais)
        MeuLV.ListView1.CheckBoxes = True
        apontaLV = 4
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If
    If ID = 6 Then '(Configurações)
        Set chamaForm = New frmConfSistema
        frmItemVerif.Show 1
        'apontaLV = 11
        'FiltroGeral = "Ativos"
        'MontaLV (apontaLV)
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
        apontaLV = 7
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 10 Then '(Fórmula - Centro de Custo)
        apontaLV = 11
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 11 Then '(FO - Ficha de Orçamento - CADASTRO INICIAL DA FCE)
        MeuLV.ListView1.CheckBoxes = True
        apontaLV = 5
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 12 Then '(Faturamento por FCE)
        apontaLV = 20
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If
    If ID = 21 Then '(FCE - Ficha de Controle de Encomenda)
        apontaLV = 6
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 26 Then '(CD - Controle de Desenhos)
        apontaLV = 10
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 27 Then '(Monitorar Produção)
        frmMonitorar.Show
    End If

'    If ID = 21 Then '(Programação)
'        MeuLV.ListView1.CheckBoxes = True
'        FiltroGeral = "Ativos pendentes"
'        apontaLV = 10
'        MontaLV (apontaLV)
'        'MeuLV.ListView1.Checkboxes = False
'    End If
    If ID = 22 Then '(LM - Lista de Materiais)
        apontaLV = 8
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 23 Then '(MP - Métodos e Processos)
        apontaLV = 9
        FiltroGeral = "Todos"
        AplicarSkin Me, Principal.Skin1
        NewColorDBGrid Me

        MontaLV (apontaLV)
    End If
    If ID = 31 Then ' Qualidade (RNCF - Registro de Não Conformidade de Fabricação)
        apontaLV = 12
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If
    If ID = 32 Then
        frmComunicacaoDesvio.Show 1
    End If
    If ID = 33 Then
        'FCRTreinCargo.Show 1
    End If

    If ID = 35 Then 'Relatórios de Inspeção
        apontaLV = 16
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 36 Then 'Impressão de Relatórios de Inspeção
        apontaLV = 19
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 41 Then ' Emissão de Relatórios de Expedição
        apontaLV = 17
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 42 Then ' Impressão de Relatórios de Expedição
        apontaLV = 18
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 51 Then '(Sistema)
        'Principal.aicAlphaImage1.Visible = True
        Set chamaForm = New frmConfSistema
        frmConfSistema.Show 1
        'Principal.aicAlphaImage1.Visible = False
    End If
    If ID = 52 Then '(Grupos)
        apontaLV = 14
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If
    If ID = 53 Then '(Usuários)
        apontaLV = 13
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If
'---------- Configurações de ambiente
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
        apontaLV = 15
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If

    If ID = 58 Then '(Terceirizados)
        apontaLV = 21
        FiltroGeral = "Todos"
        MontaLV (apontaLV)
    End If


    If ID = 71 Then '(Reabertura de OS)
        frmReabrirOP.Show
        'frmRegistro.Show 1
    End If


    If ID = 61 Then '(Sobre)
        frmRegistro.Show 1
    End If

    If ID = 62 Then '(Ajuda)
        LoadEXE (App.Path & "\ZEUSHHelp.exe")
    End If
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
'    mobjMsg.Abrir "Ajuda em construção, aguarde.", , informacao, "Master System"
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
    Dim x As Integer
    Dim nofreeze As Integer
    x = Shell(Dir, 1)
    nofreeze = DoEvents()
    Exit Sub
ERRO:
    If Err.Number = 6 Then Exit Sub
   Msgbox "Arquivo de HELP não foi localizado !!! Verifique sua localização ...", vbCritical, "Atenção"
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

        'Restaurando Permissões
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
        SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','TAB','Cadastros','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','CAT','Primários','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'02','CAT','Secundários','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0101','BUT','Ramo de atividades','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0102','BUT','Clientes','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0103','BUT','Transportadoras','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0104','BUT','Tipo material','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0205','BUT','Materiais','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0206','BUT','Itens verificação','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0207','BUT','Projetos','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'0208','BUT','Processos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'02','TAB','Orçamentos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'11','CAT','Vendas','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'1111','BUT','Serviços','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'03','TAB','Planejamento','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'21','CAT','Planejamento e Controle da Produção','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2121','BUT','FCE','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2122','BUT','LM','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2123','BUT','LD','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2124','BUT','OS','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'2125','BUT','Controle de Desenhos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'04','TAB','Produção','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'31','CAT','Acompanhamento de Produção','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'3131','BUT','OS Acompanhamento','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'3132','BUT','Evolução','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'05','TAB','Inspeção/Expedição','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'41','CAT','Emissão de Relatórios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'4141','BUT','Emitir relatório','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'4142','BUT','Imprimir relatório','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'06','TAB','Configurações','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'51','CAT','Parametrizações','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'52','CAT','Aparência','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5151','BUT','Sistema','" & vCodcoligada & "');" & _
                  "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5152','BUT','Grupos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5153','BUT','Usuários','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5254','BUT','Menu','" & vCodcoligada & "');" & _
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
                    Ribbon.AddCat right$(rsConf.Fields(2), 2), rsConf.Fields(1), rsConf.Fields(4), False
                End If
                If rsConf.Fields(3) = "BUT" Then
                    If Len(rsConf.Fields(2)) = 4 Then
                        Ribbon.AddButton right$(rsConf.Fields(2), 2), Mid$(rsConf.Fields(2), 1, 2), rsConf.Fields(4), rsConf.Fields(8)
                    Else
                        vMenu = Val(Mid$(rsConf.Fields(2), 3, 3))
                        If Len(vMenu) <> 3 Then
                            Ribbon.AddButton right$(rsConf.Fields(2), 2), Mid$(rsConf.Fields(2), 4, 2), rsConf.Fields(4), rsConf.Fields(8)
                        Else
                            Ribbon.AddButton right$(rsConf.Fields(2), 3), Mid$(rsConf.Fields(2), 3, 3), rsConf.Fields(4), rsConf.Fields(8)
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
    'Criando uma tabela temporária global
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

