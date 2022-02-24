VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesqGeral 
   BorderStyle     =   0  'None
   Caption         =   "Exemplo de Consulta usando o ListView"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   1365
   ClientWidth     =   14895
   Icon            =   "frmPesqGeral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Informações "
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   12
         Left            =   12240
         TabIndex        =   23
         Tag             =   "Afastamento/Retorno do colaborador"
         ToolTipText     =   "Afastamento/Retorno do colaborador"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":0CCA
         PICN            =   "frmPesqGeral.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   11
         Left            =   11520
         TabIndex        =   22
         Tag             =   "Atualiza Experiência"
         ToolTipText     =   "Atualiza Experiência"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":19C0
         PICN            =   "frmPesqGeral.frx":19DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Filtro "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   12960
         TabIndex        =   17
         Top             =   120
         Width           =   3135
         Begin VB.Label Label4 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   960
            TabIndex        =   21
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Período:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   10
         Left            =   10800
         TabIndex        =   16
         Tag             =   "Imprimir"
         ToolTipText     =   "Imprimir"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":26B6
         PICN            =   "frmPesqGeral.frx":26D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   8
         Left            =   10080
         TabIndex        =   14
         Tag             =   "Filtro"
         ToolTipText     =   "Filtro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":33AC
         PICN            =   "frmPesqGeral.frx":33C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   9
         Left            =   9360
         TabIndex        =   15
         Tag             =   "Admitir candidato"
         ToolTipText     =   "Admitir candidato"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":40A2
         PICN            =   "frmPesqGeral.frx":40BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   840
         Top             =   8160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":4D98
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5A72
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":674C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":7426
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":8100
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":8DDA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picBg 
         Height          =   495
         Left            =   13680
         ScaleHeight     =   435
         ScaleMode       =   0  'User
         ScaleWidth      =   936.333
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa"
         Height          =   735
         Left            =   5280
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmPesqGeral.frx":9AB4
            Left            =   120
            List            =   "frmPesqGeral.frx":9AB6
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
      End
      Begin MSComctlLib.ImageList ImgList 
         Left            =   240
         Top             =   8160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":9AB8
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":A4CA
               Key             =   "EXC"
            EndProperty
         EndProperty
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   7
         Left            =   4440
         TabIndex        =   5
         Tag             =   "Sair"
         ToolTipText     =   "Sair"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":AEDC
         PICN            =   "frmPesqGeral.frx":AEF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   6
         Left            =   3840
         TabIndex        =   6
         Tag             =   "Cancelar registro"
         ToolTipText     =   "Cancelar registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":BBD2
         PICN            =   "frmPesqGeral.frx":BBEE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   5
         Left            =   3240
         TabIndex        =   7
         Tag             =   "Editar registro"
         ToolTipText     =   "Editar registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":C8C8
         PICN            =   "frmPesqGeral.frx":C8E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   4
         Left            =   2640
         TabIndex        =   8
         Tag             =   "Novo registro"
         ToolTipText     =   "Novo registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":D5BE
         PICN            =   "frmPesqGeral.frx":D5DA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   3
         Left            =   2040
         TabIndex        =   9
         Tag             =   "Último registro"
         ToolTipText     =   "Último registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":E2B4
         PICN            =   "frmPesqGeral.frx":E2D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   2
         Left            =   1440
         TabIndex        =   10
         Tag             =   "Próximo registro"
         ToolTipText     =   "Próximo registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":EFAA
         PICN            =   "frmPesqGeral.frx":EFC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   11
         Tag             =   "Registro anterior"
         ToolTipText     =   "Registro anterior"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":FCA0
         PICN            =   "frmPesqGeral.frx":FCBC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Primeiro registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":10996
         PICN            =   "frmPesqGeral.frx":109B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7695
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList"
         SmallIcons      =   "ImgList"
         ColHdrIcons     =   "ImgList"
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmPesqGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'Acima - usado poder editar o listview --------------------

Private removeLinha As Integer
Public vStatusPDO As String
Public vDecisao As String

Private Sub Form_Load()
    IniciaBarra
    configControles
End Sub

Private Sub cmdconsulta_Click(Index As Integer)
'On Error GoTo Err
    'On Error Resume Next
    Dim Y As Integer, X As Integer
    Select Case Index
    Case 0
        Y = ListView1.ListItems.Count
        If Y > 0 Then
            ListView1.ListItems(1).Selected = True
            ListView1.ListItems(1).EnsureVisible
            ListView1.SetFocus
        End If
    Case 1
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        If X > 1 Then
            ListView1.ListItems(X - 1).Selected = True
            ListView1.ListItems(X - 1).EnsureVisible
        End If
        ListView1.SetFocus
    Case 2
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        If X < Y Then
            ListView1.ListItems(X + 1).Selected = True
            ListView1.ListItems(X + 1).EnsureVisible
        End If
        ListView1.SetFocus
    Case 3
        Y = ListView1.ListItems.Count
        If Y > 0 Then
            ListView1.ListItems(Y).Selected = True
            ListView1.ListItems(Y).EnsureVisible
            ListView1.SetFocus
        End If
    Case 4
        DesabBotoes
        Pesquisa = "novo"
        Status = "novo"
        chamaForm.Show 1
        HabBotoes
        MontaLV (apontaLV)
    Case 5
        If apontaLV = 17 Or apontaLV = 18 Then
            Unload Me
            Exit Sub
        End If
        DesabBotoes
        Pesquisa = "editar"
        AlteraListview indiceVarGlobal
        If varGlobal <> "" Then chamaForm.Show 1
        MontaLV (apontaLV)
        HabBotoes
    Case 6
        AlteraListview indiceVarGlobal
        Pesquisa = "excluir"
        CarregaSQLExcluir apontaLV
        If apontaLV <> 11 And apontaLV <> 6 And apontaLV <> 5 And apontaLV <> 4 And apontaLV <> 3 And apontaLV <> 2 And apontaLV <> 0 And apontaLV <> 16 And apontaLV <> 10 And apontaLV <> 9 And apontaLV <> 8 Then ExcluirDadosLV
        MontaLV (apontaLV)
        'gravaLog varGlobal, ListView1.SelectedItem.ListSubItems.Item(1), "-"
    Case 7
        If MeuLV.ListView1.ListItems.Count > 0 Then GravarConfLV
        Unload Me
        frmMenu2.StatusBar1.Panels(5).Text = ""
        Set chamaForm = Nothing
        Set MeuLV = Nothing
    Case 8
        FiltroGeral = ""
        TiPo = False
        DesabBotoes
        Pesquisa = "filtro"
        MontaLV (apontaLV)
        If apontaLV = 1 Then MeuLV.cmdconsulta(9).Visible = True Else MeuLV.cmdconsulta(9).Visible = False
        HabBotoes
    Case 9
        DesabBotoes
        Pesquisa = "admitir"
        AlteraListview indiceVarGlobal
        If ListView1.ListItems.Count = 0 Then
            HabBotoes
            Exit Sub
        End If
        avaliaAdmissao
    Case 10
        DesabBotoes
        Pesquisa = "Imprimir"
        If apontaLV = 9 Then
            montaTbPrintMatriz
            FCRMatrizCap.Show 1
        ElseIf apontaLV = 4 Then
            FCRListaCargos.Show 1
        ElseIf apontaLV = 0 Then
            frmPrintRels.Show 1
        ElseIf apontaLV = 18 Then
            AlteraListview indiceVarGlobal
            frmPrintRels.Show 1
        ElseIf apontaLV = 10 Then 'Programação
            'frmConvocacao.Show 1
            frmPrintRels.Show 1
        ElseIf apontaLV = 2 Or apontaLV = 3 Or apontaLV = 5 Or apontaLV = 6 Or apontaLV = 11 Or apontaLV = 17 Then
            FCRGeral.Show 1
        ElseIf apontaLV = 16 Then
            frmPrintRels.Show 1
        End If
        HabBotoes
    Case 11
        caculaTmpExp
        MontaLV (apontaLV)
    Case 12
        AlteraListview 1
        If varGlobal <> "" Then afastaColaborador
        
        FiltroGeral = ""
        TiPo = False
        DesabBotoes
        Pesquisa = "filtro"
        MontaLV (apontaLV)
        If apontaLV = 1 Then cmdconsulta(9).Visible = True Else cmdconsulta(9).Visible = False
        HabBotoes
        frmMenu2.StatusBar1.Panels(5).Text = ""
        
    End Select
    configControles
    Exit Sub
Err:
    MsgBox "Nenhum item selecionado", vbInformation, "SGCH"
    Exit Sub
End Sub

Private Sub avaliaAdmissao()
'-Padrao - para saber se ja tem uma solicitação cadastrada --------------------------------
    Dim vNumPDO As Integer
    Dim rsPDOColab As New ADODB.Recordset
    Dim SqlPDOColab As String
   
    SqlPDOColab = "Select a.cpf,a.codcolaborador,a.nomecolaborador,b.id,b.status,b.tipo,b.decisao,a.datarecisao from tbcolaboradores as a left join tbautorizacao as b on a.autorizacao = b.id where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & Mid$(varGlobal, 1, 11) & "' and a.datarecisao is null"
    rsPDOColab.Open SqlPDOColab, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not IsNull(rsPDOColab.Fields(7)) Then
        MsgBox "Colaborador DEMITIDO, não pode ser admitido através desse módulo"
        HabBotoes
        Exit Sub
    End If
    
    If Not IsNull(rsPDOColab.Fields(3)) Then
        If rsPDOColab.RecordCount > 0 Then
            vNumPDO = rsPDOColab.Fields(3)
            If rsPDOColab.Fields(4) = "N" Or IsNull(rsPDOColab.Fields(4)) Then
                MsgBox "O PDO nº: " & Format(vNumPDO, "000000") & " esta em aberto para este " & rsPDOColab.Fields(5) & ". Aguarde tomada de decisão", vbCritical, "Atenção"
                rsPDOColab.Close
                Set rsPDOColab = Nothing
                HabBotoes
                Exit Sub
            Else
                If Not IsNull(rsPDOColab.Fields(4)) Then
                    vStatusPDO = rsPDOColab.Fields(4)
                    vDecisao = rsPDOColab.Fields(6)
                End If
            End If
        End If
    End If
    rsPDOColab.Close
    Set rsPDOColab = Nothing
    
    If apontaLV = 0 Then
        If vStatusPDO <> "S" Then
            If ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK" Then
                MsgBox "Colaborador já admitido", vbInformation, "SGCH"
                HabBotoes
                Exit Sub
            End If
            If Val(RemoveMask(ListView1.SelectedItem.ListSubItems.Item(5))) < MediaGlobal And Val(RemoveMask(ListView1.SelectedItem.ListSubItems.Item(5))) >= vAprovadoRest Then
                If vAdiRes = "N" Then
                    If MsgBox("Pontuação do colaborador está abaixo da média. Usúario não privilégios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                        gravaSolicitacao Mid$(varGlobal, 1, 11), "colaborador", RemoveMask(ListView1.SelectedItem.ListSubItems.Item(5)), "A pontuação do colaborador está abaixo da média parametrizada para Aprovação no sistema para o cargo: " & ListView1.SelectedItem.ListSubItems.Item(8), NomUsu
                        MsgBox "Foi gerado o PDO nº: " & Format(vPDO, "000000") & ". Aguarde tomada de decisão", vbInformation, "SGCH"
                    End If
                    configControles
                    HabBotoes
                    Exit Sub
                End If
            End If
            If Val(RemoveMask(ListView1.SelectedItem.ListSubItems.Item(5))) < vAprovadoRest Then
                If vAdiRep = "N" Then
                    If MsgBox("Pontuação do colaborador está abaixo da média. Usúario não privilégios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                        gravaSolicitacao Mid$(varGlobal, 1, 11), "colaborador", RemoveMask(ListView1.SelectedItem.ListSubItems.Item(5)), "A pontuação do colaborador está abaixo da média parametrizada para Aprovação com Restrição no sistema para o cargo: " & ListView1.SelectedItem.ListSubItems.Item(8), NomUsu
                        MsgBox "Foi gerado o PDO nº: " & Format(vPDO, "000000") & ". Aguarde tomada de decisão", vbInformation, "SGCH"
                    End If
                    configControles
                    HabBotoes
                    Exit Sub
                End If
            End If
        
        End If
    End If
    
    If apontaLV = 1 Then
        If vStatusPDO <> "S" Then
            If Val(RemoveMask(ListView1.SelectedItem.ListSubItems.Item(4))) < MediaGlobal And Val(RemoveMask(ListView1.SelectedItem.ListSubItems.Item(4))) >= vAprovadoRest Then
                If vAdiRes = "N" Then
                    If MsgBox("Pontuação do colaborador está abaixo da média. Usúario não privilégios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                        gravaSolicitacao Mid$(varGlobal, 1, 11), "colaborador", RemoveMask(ListView1.SelectedItem.ListSubItems.Item(4)), "A pontuação do colaborador está abaixo da média parametrizada para Aprovação no sistema para o cargo: " & ListView1.SelectedItem.ListSubItems.Item(7), NomUsu
                        MsgBox "Foi gerado o PDO nº: " & Format(vPDO, "000000") & ". Aguarde tomada de decisão", vbInformation, "SGCH"
                    End If
                    configControles
                    HabBotoes
                    Exit Sub
                End If
            End If
        
            If Val(RemoveMask(ListView1.SelectedItem.ListSubItems.Item(4))) < vAprovadoRest Then
                If vAdiRep = "N" Then
                    If MsgBox("Pontuação do colaborador está abaixo da média. Usúario não privilégios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                        gravaSolicitacao Mid$(varGlobal, 1, 11), "colaborador", RemoveMask(ListView1.SelectedItem.ListSubItems.Item(4)), "A pontuação do colaborador está abaixo da média parametrizada para Aprovação com Restrição no sistema para o cargo: " & ListView1.SelectedItem.ListSubItems.Item(7), NomUsu
                        MsgBox "Foi gerado o PDO nº: " & Format(vPDO, "000000") & ". Aguarde tomada de decisão", vbInformation, "SGCH"
                    End If
                    configControles
                    HabBotoes
                    Exit Sub
                End If
            End If
        End If
    End If
    If varGlobal <> "" Then frmAdmitirCandidato.Show 1
    HabBotoes
End Sub

Private Sub montaTbPrintMatriz()
    Dim rsMatriz As New ADODB.Recordset
    Dim SqlMatriz As String
    
    Dim rsPrintMatriz As New ADODB.Recordset
    Dim SqlPrintMatriz As String
    Dim vCodMatriz As Integer
    
    cnBanco.BeginTrans
    
    SqlPrintMatriz = "Delete from tbPrintMatriz where codcoligada = '" & vCodcoligada & "'"
    rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
    SqlMatriz = "select * from tbMatriz where codcoligada = '" & vCodcoligada & "' order by codmatriz"
    rsMatriz.Open SqlMatriz, cnBanco, adOpenKeyset, adLockOptimistic
    While Not rsMatriz.EOF
        vCodMatriz = rsMatriz.Fields(0)

        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "ESCOLARIDADE" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo5,codcoligada) Values(" & vCodMatriz & ",'ESCOLARIDADE','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "ESCOLARIDADE" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo2,campo3,campo4,codcoligada) Select a.codmatriz,a.codescolaridade,b.nomeescolaridade,str(a.pontuacao)+'%',a.codcoligada from tbmatrizesc as a inner join tbescolaridade as b on b.codescolaridade = a.codescolaridade where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & vCodMatriz & "'"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "CURSOS/TREINAMENTOS" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo5,codcoligada) Values(" & vCodMatriz & ",'CURSOS/TREINAMENTOS','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
        ''ABAIXO - Insere a "CURSOS/TREINAMENTOS" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo2,campo3,codcoligada) Select a.codmatriz,a.codtreinamento,b.nometreinamento,a.codcoligada from tbMatrizCur as a, tbTreinamentos as b where a.codcoligada = '" & vCodcoligada & "' and b.codtreinamento = a.codtreinamento and a.codmatriz = '" & vCodMatriz & "'"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "HABILIDADES" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo5,codcoligada) Values(" & vCodMatriz & ",'HABILIDADES','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "HABILIDADES" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo2,campo3,codcoligada) Select a.codmatriz,a.codhabilidade,b.nomehabilidade,a.codcoligada from tbMatrizHab as a, tbHabilidades as b Where a.codcoligada = '" & vCodcoligada & "' and b.codhabilidade = a.codhabilidade and a.codmatriz = '" & vCodMatriz & "'"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
        '******************************************************************************
        '*** ABAIXO - Insere o nome da Competência "EXPERIÊNCIA" na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo5,codcoligada) Values(" & vCodMatriz & ",'EXPERIÊNCIA','" & vCodcoligada & "')"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco

        ''ABAIXO - Insere a "EXPERIÊNCIA" referente a MATRIZ selecionada na tabela tbPrintMatriz
        SqlPrintMatriz = "Insert into tbPrintMatriz(campo1,campo2,campo3,campo4,codcoligada) Select a.codmatriz,a.codcargo,b.nomecargo,a.tmpoexp,a.codcoligada from tbmatrizexp as a, tbcargos as b where a.codcoligada = '" & vCodcoligada & "' and b.codcargo = a.codcargo and a.codmatriz = '" & vCodMatriz & "'"
        rsPrintMatriz.Open SqlPrintMatriz, cnBanco
    
        '******************************************************************************
        rsMatriz.MoveNext
    Wend
    cnBanco.CommitTrans
End Sub

Private Sub DesabBotoes()
On Error Resume Next
    Dim X As Integer
    For X = 0 To MeuLV.cmdconsulta.Count - 1
        If MeuLV.cmdconsulta(X).Visible = True Then MeuLV.cmdconsulta(X).UseGreyscale = True
    Next
    If vIntegra = "S" Then
        MeuLV.cmdconsulta(6).UseGreyscale = True
        MeuLV.cmdconsulta(6).DragMode = 1
        MeuLV.cmdconsulta(6).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub HabBotoes()
On Error Resume Next
    Dim X As Integer
    For X = 0 To MeuLV.cmdconsulta.Count - 1
        If MeuLV.cmdconsulta(X).Visible = True Then MeuLV.cmdconsulta(X).UseGreyscale = False
    Next
    If vIntegra = "S" Then
        MeuLV.cmdconsulta(6).UseGreyscale = True
        MeuLV.cmdconsulta(6).DragMode = 1
        MeuLV.cmdconsulta(6).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub AlteraListview(qtdCol As Integer)
    On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            If ListView1.CheckBoxes = True Then ListView1.ListItems.Item(X).Checked = True
            Exit For
        End If
    Next
    If qtdCol = 1 Then
        varGlobal = ListView1.ListItems.Item(X)
    Else
        varGlobal = ListView1.ListItems.Item(X) & ListView1.SelectedItem.ListSubItems.Item(1)
    End If
    removeLinha = X
    Exit Sub
Err:
    varGlobal = ""
    MsgBox "Nenhum Cargo cadastrado ou selecionado", vbInformation, "SGCH"
    Exit Sub
End Sub

Private Sub Pesquisar(Optional Column As ColumnHeader = Nothing)
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count 'Conta as linhas preenchidas do Listview
    If Y > 0 Then 'Entra nessa condição se o Listview não estiver vazio
        Dim c As ColumnHeader
        Dim numCol As Integer
        numCol = 0
        For Each c In MeuLV.ListView1.ColumnHeaders
            If Combo1.Text = c Then Exit For
            numCol = numCol + 1
        Next
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True 'Seleciona a linha de acordo com o valor de "X"
            'SE FOR SELECIONADO A PRIMEIRA COLUNA
            If Combo1.Text = "" Then
                'Se não for selecionado nada no ComboBox Combo1
                MsgBox "Nenhum filtro de pesquisa selecionado"
                Exit Sub
            End If
            If numCol = 0 Then
                If UCase(ListView1.ListItems.Item(X)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    ListView1.SetFocus
                    Exit Sub
                End If
            'SE FOR SELECIONADO A PARTIR DA SEGUNDA COLUNA
            ElseIf numCol > 0 Then
                If UCase(ListView1.SelectedItem.ListSubItems.Item(numCol)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    ListView1.SetFocus
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub IniciaBarra()
    '-------------------------
    'Incializa o estilo do PictureBox
    '------------------------
    picBg.BackColor = ListView1.BackColor
    picBg.ScaleMode = vbTwips
    picBg.BorderStyle = vbBSNone
    picBg.AutoRedraw = True
    picBg.Visible = False
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        DesabBotoes
        Pesquisa = "editar"
        AlteraListview indiceVarGlobal
        If varGlobal <> "" Then chamaForm.Show 1
        HabBotoes
        configControles
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        Pesquisar ' Sub que realiza a Pesquisa no Listview mediante ao que foi digitado no TexBox Text1 e ao q foi selecionado no ComboBox Combo1
    End If
End Sub

'As duas Subs abaixo faz com que ordene o listview pela coluna que vc clicar
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView1, ColumnHeader
    Combo1.Text = ColumnHeader.Text
End Sub

Public Sub ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
    With ListView1
    If .SortKey <> Column.Index - 1 Then
        .SortKey = Column.Index - 1
        .SortOrder = lvwAscending
    Else
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End If
    .Sorted = -1
    End With
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdconsulta(4).UseGreyscale = True
        cmdconsulta(4).DragMode = 1
        cmdconsulta(4).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdconsulta(5).UseGreyscale = True
        cmdconsulta(5).DragMode = 1
        cmdconsulta(5).SpecialEffect = cbEngraved
    End If
    'If vSal = "N" Then
        'cmdconsulta(4).UseGreyscale = True
    'End If
    If vExc = "N" Then
        cmdconsulta(6).UseGreyscale = True
        cmdconsulta(6).DragMode = 1
        cmdconsulta(6).SpecialEffect = cbEngraved
    End If
    If vImp = "N" Then
        cmdconsulta(10).UseGreyscale = True
        cmdconsulta(10).DragMode = 1
        cmdconsulta(10).SpecialEffect = cbEngraved
    End If
    If vFil = "N" Then
        cmdconsulta(8).UseGreyscale = True
        cmdconsulta(8).DragMode = 1
        cmdconsulta(8).SpecialEffect = cbEngraved
    End If
    'If vAva = "N" Then
        'cmdconsulta(4).UseGreyscale = True
    'End If
    If vAdi = "N" Then
        cmdconsulta(9).UseGreyscale = True
        cmdconsulta(9).DragMode = 1
        cmdconsulta(9).SpecialEffect = cbEngraved
    End If
    'If vDem = "N" Then
        'cmdconsulta(4).UseGreyscale = True
    'End If
End Sub



'**********************************************
'**********************************************
'**********************************************
'**********************************************
'**********************************************

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If apontaLV = 4 Then
        Dim i As Integer, leftPos As Single 'the left pos of the column
        Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate
        If Button = vbLeftButton Then
            If Not ListView1.SelectedItem Is Nothing Then
                ListView1.LabelEdit = lvwManual
                dx = GetLvwDeltaX
                lvwX = X + dx
                For i = 4 To 4
                    leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
                    If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                        m_RowIndex = ListView1.SelectedItem.Index 'row
                        m_ColIndex = i 'column
                            AlteraListview indiceVarGlobal
                            If varGlobal <> "" Then AtivaDesativaCago
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
End Sub

Function GetLvwDeltaX() As Single
    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView1.ColumnHeaders(ListView1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView1.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub AtivaDesativaCago()
    Dim rsAtivaDesativaCago As New ADODB.Recordset
    Dim SqlAtivaDesativaCago As String
    Dim vAtivo As String
    SqlAtivaDesativaCago = "update tbcargos set ativo = case  WHEN codcoligada = '" & vCodcoligada & "' and ativo = 'S' then 'N' WHEN codcoligada = '" & vCodcoligada & "' and ativo = 'N' then 'S' ELSE 'S' END where codcoligada = '" & vCodcoligada & "' and codcargo = '" & Val(varGlobal) & "'"
    rsAtivaDesativaCago.Open SqlAtivaDesativaCago, cnBanco
    
    SqlAtivaDesativaCago = "Select ativo from tbcargos where codcoligada = '" & vCodcoligada & "' and codcargo = '" & Val(varGlobal) & "'"
    rsAtivaDesativaCago.Open SqlAtivaDesativaCago, cnBanco, adOpenKeyset, adLockReadOnly
    vAtivo = rsAtivaDesativaCago.Fields(0)
    If vAtivo <> "S" Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "EXC"
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "OK"
    End If
    rsAtivaDesativaCago.Close
End Sub
