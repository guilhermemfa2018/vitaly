VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesqGeral 
   BorderStyle     =   0  'None
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   1365
   ClientWidth     =   17055
   Icon            =   "frmPesqGeral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   17055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Informações "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16695
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   0
         Left            =   11040
         TabIndex        =   21
         Tag             =   "Plano de Programação Semanal"
         ToolTipText     =   "Plano de Programação Semanal"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
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
            NumListImages   =   29
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":19C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":269A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3374
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":404E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":4D28
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5A02
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":66DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":73B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":8090
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":8D6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":9A44
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":A71E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":B3F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":C0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":CDAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":DA86
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":E760
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":F43A
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":10114
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":10DEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":11AC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":127A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":1347C
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":14156
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":14E30
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":15B0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":167E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":174BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":18198
               Key             =   ""
            EndProperty
         EndProperty
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
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":18E72
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":19884
               Key             =   "EXC"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":1A296
               Key             =   "POSITIVO"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":1ECB0
               Key             =   "NEGATIVO"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":236CA
               Key             =   "ARQUIVADO"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":31A53
               Key             =   "AGUARDE-01"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3272D
               Key             =   "AGUARDE-02"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":33407
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":340E1
               Key             =   "PENDENTE12"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":34DBB
               Key             =   "AVALIANDO1"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":35A95
               Key             =   "CONCLUIDO1"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3676F
               Key             =   "PRETO"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":37449
               Key             =   "PENDENTE1"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":38123
               Key             =   "FABRICANDO"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":38DFD
               Key             =   "FECHADO"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":39AD7
               Key             =   "ANDAMENTO"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3A7B1
               Key             =   "CONCLUIDA"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3B1C3
               Key             =   "PARALIZADA"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3BE9D
               Key             =   "DUVIDA"
            EndProperty
         EndProperty
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   12
         Left            =   10440
         TabIndex        =   15
         Tag             =   "Afastamento/Retorno do colaborador"
         ToolTipText     =   "Afastamento/Retorno do colaborador"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":3CB77
         PICN            =   "frmPesqGeral.frx":3CB93
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   11
         Left            =   9840
         TabIndex        =   14
         Tag             =   "Atualiza Experiência"
         ToolTipText     =   "Atualiza Experiência"
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":3D86D
         PICN            =   "frmPesqGeral.frx":3D889
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   10
         Left            =   9240
         TabIndex        =   13
         Tag             =   "Imprimir"
         ToolTipText     =   "Imprimir"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":3E563
         PICN            =   "frmPesqGeral.frx":3E57F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   8
         Left            =   8640
         TabIndex        =   12
         Tag             =   "Filtro"
         ToolTipText     =   "Filtro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":3F259
         PICN            =   "frmPesqGeral.frx":3F275
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   9
         Left            =   8040
         TabIndex        =   11
         Tag             =   "Admitir candidato"
         ToolTipText     =   "Admitir candidato"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":3FF4F
         PICN            =   "frmPesqGeral.frx":3FF6B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   7
         Left            =   1920
         TabIndex        =   10
         Tag             =   "Sair"
         ToolTipText     =   "Sair"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":40C45
         PICN            =   "frmPesqGeral.frx":40C61
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   6
         Left            =   1320
         TabIndex        =   9
         Tag             =   "Cancelar registro"
         ToolTipText     =   "Cancelar registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":4193B
         PICN            =   "frmPesqGeral.frx":41957
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   5
         Left            =   720
         TabIndex        =   8
         Tag             =   "Editar registro"
         ToolTipText     =   "Editar registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":42631
         PICN            =   "frmPesqGeral.frx":4264D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Tag             =   "Novo registro"
         ToolTipText     =   "Novo registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   11
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":43327
         PICN            =   "frmPesqGeral.frx":43343
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
         Left            =   12360
         TabIndex        =   6
         Top             =   120
         Width           =   3975
         Begin ACTIVESKINLibCtl.SkinLabel Label3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmPesqGeral.frx":4401D
            TabIndex        =   19
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label4 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmPesqGeral.frx":44085
            TabIndex        =   18
            Top             =   480
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label2 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmPesqGeral.frx":440DF
            TabIndex        =   17
            Top             =   240
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmPesqGeral.frx":44139
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picBg 
         Height          =   495
         Left            =   15600
         ScaleHeight     =   435
         ScaleMode       =   0  'User
         ScaleWidth      =   936.333
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   5175
         Begin ZEUS.chameleonButton chameleonButton1 
            Height          =   495
            Left            =   4560
            TabIndex        =   20
            Tag             =   "Pesquisar"
            ToolTipText     =   "Pesquisar"
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            BTYPE           =   11
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmPesqGeral.frx":4419F
            PICN            =   "frmPesqGeral.frx":441BB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            Left            =   2400
            TabIndex        =   3
            Top             =   240
            Width           =   2055
         End
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
            ItemData        =   "frmPesqGeral.frx":44935
            Left            =   120
            List            =   "frmPesqGeral.frx":44937
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7695
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   16455
         _ExtentX        =   29025
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
         ForeColor       =   8388608
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
Public vDecisao As String, vRetrabalho As String
Public vX As Integer, vY As Integer, vPosAtual As Integer

Private Sub chameleonButton1_Click()
    Pesquisar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    vPosAtual = 1
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    configControles
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub cmdconsulta_Click(Index As Integer)
'On Error GoTo Err
    'On Error Resume Next
    Dim Y As Integer, X As Integer
    Select Case Index
    Case 0
        
        If apontaLV = 9 Then
            'AlteraListview 1
            frmProgramacao.Show
        ElseIf apontaLV = 20 Then
            AlteraListview 2
            varGlobal = Mid$(varGlobal, 1, 4)
            frmAlteraStatusFCE.Show 1
        End If
        
        'y = ListView1.ListItems.Count
        'If y > 0 Then
        '    ListView1.ListItems(1).Selected = True
        '    ListView1.ListItems(1).EnsureVisible
        '    ListView1.SetFocus
        'End If
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
        vTime = Time
        vTime = RemoveMask(vTime)
        If apontaLV = 16 Then
            AlteraListview 2
        Else
            AlteraListview 1
        End If
        If varGlobal2 = "?" Then
            'Me.MousePointer = 0
            varGlobal2 = ""
            Exit Sub
        End If
        
        DesabBotoes
        Pesquisa = "novo"
        Status = "novo"
        If apontaLV = 6 Then
            AlteraListview indiceVarGlobal
            frmLM.Show 1
        Else
            'Me.MousePointer = 0
            If apontaLV = 16 Or apontaLV = 17 Then
                If apontaLV = 16 Then
                    vSituacao = "INSPEÇÃO DE FABRICAÇÃO"
                Else
                    vSituacao = "EXPEDIÇÃO"
                    Set chamaForm = New frmRelExp
                End If
                AlteraListview 2
                chamaForm.Show
                Exit Sub
            Else
                chamaForm.Show 1
            End If
        End If
        HabBotoes
        MontaLV (apontaLV)
    Case 5
        vTime = Time
        vTime = RemoveMask(vTime)
        AlteraListview 1
        
        If varGlobal2 = "?" Then
            'Me.MousePointer = 0
            mobjMsg.Abrir "FCE concluida. A LM não pode ser EDITADA", Ok, critico, "Atenção"
            varGlobal2 = ""
            Exit Sub
        End If
        
        If apontaLV = 18 Then
            Unload Me
            Exit Sub
        End If
        DesabBotoes
        Pesquisa = "editar"
        AlteraListview indiceVarGlobal
        If varGlobal <> "" Then
            If apontaLV = 9 And vRetrabalho <> "-" Then
                frmRetrabalho.Show 1
            ElseIf apontaLV = 17 Then
                'Me.MousePointer = 0
                vSituacao = "EXPEDIÇÃO TERC."
                frmRelExpAvulso.Show
                Exit Sub
            Else
                chamaForm.Show 1
                Exit Sub
            End If
        End If
        MontaLV (apontaLV)
        HabBotoes
    Case 6
        If apontaLV <> 9 And apontaLV <> 16 Then
            AlteraListview indiceVarGlobal
        Else
            AlteraListview 2
        End If
        If apontaLV = 16 Then
            vSituacao = "INSPEÇÃO DE PINTURA"
            chamaForm.Show
            Exit Sub
        End If
        
        Pesquisa = "excluir"
        CarregaSQLExcluir apontaLV
        If apontaLV <> 11 And apontaLV <> 6 And apontaLV <> 5 And apontaLV <> 4 And apontaLV <> 3 And apontaLV <> 2 And apontaLV <> 0 And apontaLV <> 10 And apontaLV <> 9 And apontaLV <> 8 And apontaLV <> 15 And apontaLV <> 18 And apontaLV <> 19 Then ExcluirDadosLV
        MontaLV (apontaLV)
        'gravaLog varGlobal, ListView1.SelectedItem.ListSubItems.Item(1), "-"
    Case 7
        If MeuLV.ListView1.ListItems.Count > 0 Then GravarConfLV
        Principal.StatusBar1.Panels(5).Text = "Registros: "
        Unload Me
        Set chamaForm = Nothing
        Set MeuLV = Nothing
    Case 8
        FiltroGeral = ""
        Tipo = False
        DesabBotoes
        Pesquisa = "filtro"
        MontaLV (apontaLV)
        'If apontaLV = 9 Or apontaLV = 12 Then cmdconsulta(9).Visible = True Else cmdconsulta(9).Visible = False
        HabBotoes
        Principal.StatusBar1.Panels(5).Text = "Registros: " & MeuLV.ListView1.ListItems.Count
    Case 9
        If apontaLV = 9 Then
            'AlteraListview 1
            frmComunicacaoDesvio.Show 1
        ElseIf apontaLV = 5 Then
            frmRecFO.Show 1
        ElseIf apontaLV = 12 Then
            AlteraListview 1
            frmCausais.Show 1
        End If
    Case 10
        DesabBotoes
        Pesquisa = "Imprimir"
        If apontaLV = 9 Or apontaLV = 12 Then
            frmPrintRels.Show 1
            'FCRConfronto.Show 1
        ElseIf apontaLV = 4 Then
            'FCRListaCargos.Show 1
        ElseIf apontaLV = 0 Then
            'frmPrintRels.Show 1
        ElseIf apontaLV = 18 Then
            'AlteraListview indiceVarGlobal
            'frmPrintRels.Show 1
        ElseIf apontaLV = 19 Then
            frmPrintRels.Show 1
        ElseIf apontaLV = 20 Then 'Comercial - Faturamento
            AlteraListview indiceVarGlobal
            'If ListView1.SelectedItem.ListSubItems.Item(3) = "-" Then
            '    mobjMsg.Abrir "Não Exitem lançamentos financeiros para a FCE nº: " & Mid$(varGlobal, 1, 4), Ok, critico, "ZEUS"
            '    Exit Sub
            'End If
            montaDadosVendas
            frmPrintRels.Show 1
        ElseIf apontaLV = 2 Or apontaLV = 3 Or apontaLV = 5 Or apontaLV = 6 Or apontaLV = 11 Or apontaLV = 17 Then
            'FCRGeral.Show 1
        End If
        HabBotoes
    Case 11
        If apontaLV = 9 Then
            AlteraListview 1
            
            If vRetrabalho <> "-" Then
                Pesquisa = "editar"
            Else
                Pesquisa = "novo"
            End If
            vTime = Time
            vTime = RemoveMask(vTime)
            frmRetrabalho.Show 1
        ElseIf apontaLV = 5 Then
            AlteraListview 1
            If varGlobal2 <> "-" And varGlobal2 <> "?" Then
                frmFCE.Show 1
            ElseIf varGlobal2 = "?" Then
                varGlobal2 = ""
                Exit Sub
            Else
                mobjMsg.Abrir "Nenhuma FCE selecionada", Ok, critico, "ZEUS"
            End If
        End If
        'caculaTmpExp
        'MontaLV (apontaLV)
    Case 12
        If apontaLV = 9 Then
            AlteraListview 2
            frmBaixaParcialOS.Show 1
            'mobjMsg.Abrir "Rotina de Baixa parcial de OS em desenvolvimento", Ok, informacao, "Atenção"
        End If
'        AlteraListview 1
'        FiltroGeral = ""
'        Tipo = False
'        DesabBotoes
'        Pesquisa = "filtro"
'        MontaLV (apontaLV)
'        If apontaLV = 1 Then cmdconsulta(9).Visible = True Else cmdconsulta(9).Visible = False
'        HabBotoes
'        Principal.StatusBar1.Panels(5).Text = ""
        
    End Select
    configControles
    Exit Sub
Err:
    mobjMsg.Abrir "Nenhum item selecionado", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub VerificaFCE()
    On Error GoTo Err
    Dim Y As Integer, X As Integer
    Dim fce As String
    Y = ListView1.ListItems.Count
    fce = ""
    For X = 1 To Y
        ListView1.ListItems(X).Selected = True
        If ListView1.ListItems.Item(X).Selected = True Then
            If ListView1.ListItems.Item(X).Checked = True Then
                varGlobal = ListView1.ListItems.Item(X)
                If fce <> "" Then
                    If ListView1.SelectedItem.ListSubItems.Item(1) <> fce Then
                        mobjMsg.Abrir "Não é permitido selecionadas FO's de empresas diferentes", Ok, critico, "Atenção"
                        Contador = 0
                        Exit Sub
                    End If
                End If
                
                fce = ListView1.SelectedItem.ListSubItems.Item(1)
                If ListView1.SelectedItem.ListSubItems.Item(13) <> "" And ListView1.SelectedItem.ListSubItems.Item(13) <> "-" Then
                    mobjMsg.Abrir "A FO selecionada ja esta em uma FCE", Ok, critico, "Atenção"
                    Contador = 0
                    Exit Sub
                Else
                    Contador = Contador + 1
                End If
            End If
        End If
    Next
    Exit Sub
Err:
    mobjMsg.Abrir "Nenhuma Ficha de Orçamento selecionada", Ok, critico, "Atenção"
    Exit Sub
End Sub


Private Sub DesabBotoes()
On Error Resume Next
    Dim X As Integer
'    For X = 0 To cmdconsulta.Count - 1
'        If cmdconsulta(X).Visible = True Then cmdconsulta(X).UseGreyscale = True
'    Next
'    If vIntegra = "S" Then
'        cmdconsulta(6).UseGreyscale = True
'        cmdconsulta(6).DragMode = 1
'        cmdconsulta(6).SpecialEffect = cbEngraved
'    End If
End Sub

Private Sub HabBotoes()
On Error Resume Next
'    Dim X As Integer
'    For X = 0 To cmdconsulta.Count - 1
'        If cmdconsulta(X).Visible = True Then cmdconsulta(X).UseGreyscale = False
'    Next
'    If vIntegra = "S" Then
'        cmdconsulta(6).UseGreyscale = True
'        cmdconsulta(6).DragMode = 1
'        cmdconsulta(6).SpecialEffect = cbEngraved
'    End If
End Sub

Private Sub AlteraListview(qtdCol As Integer)
    On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            'If ListView1.CheckBoxes = True Then ListView1.ListItems.Item(X).Checked = True
            Exit For
        End If
    Next
    If qtdCol = 1 Then
        varGlobal = ListView1.ListItems.Item(X)
    ElseIf qtdCol = 3 Then
        varGlobal = ListView1.SelectedItem.ListSubItems.Item(1)
    Else
        varGlobal = ListView1.ListItems.Item(X) & ListView1.SelectedItem.ListSubItems.Item(1)
    End If
    If apontaLV = 9 Then
        vRetrabalho = ListView1.SelectedItem.ListSubItems.Item(9)
    End If
    If apontaLV = 5 Then
        If ListView1.SelectedItem.ListSubItems.Item(16).ReportIcon = "CONCLUIDA" Then
            varGlobal2 = "?"
            mobjMsg.Abrir "FCE concluida não pode ser editada", Ok, critico, "Atenção"
        End If
        If varGlobal2 <> "?" Then varGlobal2 = ListView1.SelectedItem.ListSubItems.Item(13)
    ElseIf apontaLV = 6 Then
        'SE A FCE ESTIVER CONCLUIDA NÃO CONSEGUE CRIAR NOVAS LISTAS DE MATERIAIS
        If ListView1.SelectedItem.ListSubItems.Item(12).ReportIcon = "CONCLUIDA" Then
            varGlobal2 = "?"
            'mobjMsg.Abrir "FCE concluida não pode ser criado nova LM", Ok, critico, "Atenção"
        End If
    ElseIf apontaLV = 8 Then
        'SE A FCE ESTIVER CONCLUIDA NÃO CONSEGUE EDITAR A LISTA DE MATERIAIS
        If ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "CONCLUIDA" Then
            varGlobal2 = "?"
            'mobjMsg.Abrir "FCE concluida. A LM não pode ser EDITADA", Ok, critico, "Atenção"
        End If
    End If
    
    If apontaLV = 18 Then vQualquerDado(20, 1) = ListView1.SelectedItem.ListSubItems.Item(1)
    
    
    If apontaLV = 19 Then
        'SOMENTE PARA OS RELATÓRIOS DE INSPEÇÃO DE PINTURA
        Dim rsInspecao As New ADODB.Recordset
        Dim sqlInspecao As String
        limpaQualquerDado
        
        sqlInspecao = "select b.descricao,b.sigla from tbVerifGrupo as a inner join tbVerifItem as b on a.codgrupo = b.codgrupo where a.aplicacao <> '-'"
        rsInspecao.Open sqlInspecao, cnBanco, adOpenKeyset, adLockReadOnly
        Y = rsInspecao.RecordCount
        For X = 1 To Y
            vQualquerDado(0, X) = rsInspecao.Fields(1) & " - " & rsInspecao.Fields(0)
            rsInspecao.MoveNext
        Next
        rsInspecao.Close
        Set rsInspecao = Nothing
        vQualquerDado(0, 30) = "RELATÓRIO DE INSPEÇÃO DE " & ListView1.SelectedItem.ListSubItems.Item(6)
    End If
    
    removeLinha = X
    Exit Sub
Err:
    varGlobal = ""
    If ListView1.ListItems.Count <> 0 Then
        mobjMsg.Abrir "Nenhum registro cadastrado ou selecionado", Ok, critico, "Atenção"
    End If
    Exit Sub
End Sub

Private Sub Pesquisar(Optional Column As ColumnHeader = Nothing)
    vY = ListView1.ListItems.Count 'Conta as linhas preenchidas do Listview
    If vY > 0 Then 'Entra nessa condição se o Listview não estiver vazio
        Dim c As ColumnHeader
        Dim numCol As Integer
        numCol = 0
        For Each c In ListView1.ColumnHeaders
            If Combo1.Text = c Then Exit For
            numCol = numCol + 1
        Next
        For vX = vPosAtual To vY
            ListView1.ListItems(vX).Selected = True 'Seleciona a linha de acordo com o valor de "X"
            'SE FOR SELECIONADO A PRIMEIRA COLUNA
            If Combo1.Text = "" Then
                'Se não for selecionado nada no ComboBox Combo1
                mobjMsg.Abrir "Nenhum filtro de pesquisa selecionado", Ok, critico, "Atenção"
                vPosAtual = vX + 1
                Exit Sub
            End If
            If numCol = 0 Then
                If UCase(ListView1.ListItems.Item(vX)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(vX).Selected = True
                    ListView1.ListItems(vX).EnsureVisible
                    ListView1.SetFocus
                    vPosAtual = vX
                    Exit Sub
                End If
            'SE FOR SELECIONADO A PARTIR DA SEGUNDA COLUNA
            ElseIf numCol > 0 Then
                If UCase(ListView1.SelectedItem.ListSubItems.Item(numCol)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(vX).Selected = True
                    ListView1.ListItems(vX).EnsureVisible
                    ListView1.SetFocus
                    vPosAtual = vX + 1
                    Exit Sub
                End If
            End If
            If vX >= vY Then
                vPosAtual = 1
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

Private Sub Form_Resize()
'    OrganizaControles
End Sub

Private Sub ListView1_DblClick()
    Me.MousePointer = 11
    If vEdi <> "N" Then
        'DesabBotoes
        Pesquisa = "editar"
        If apontaLV = 16 Or apontaLV = 17 Then
            If apontaLV = 16 Then
                vSituacao = "INSPEÇÃO DE FABRICAÇÃO"
            Else
                vSituacao = "EXPEDIÇÃO"
            End If
            AlteraListview 2
        Else
            AlteraListview indiceVarGlobal
        End If
        If varGlobal2 = "?" Then
            mobjMsg.Abrir "FCE concluida. A LM não pode ser EDITADA", Ok, critico, "Atenção"
            varGlobal2 = ""
            Exit Sub
        End If
        
        If varGlobal <> "" Then
            vTime = Time
            vTime = RemoveMask(vTime)
            If apontaLV = 9 And vRetrabalho <> "-" Then
                frmRetrabalho.Show 1
            Else
                If apontaLV = 16 Or apontaLV = 17 Then
                    chamaForm.Show
                Else
                    If apontaLV = 20 Then
                        If ListView1.SelectedItem.ListSubItems.Item(3) = "-" Then
                            mobjMsg.Abrir "Não Exitem lançamentos financeiros para a FCE nº: " & Mid$(varGlobal, 1, 4), Ok, critico, "ZEUS"
                            'rsFatFCE.Close
                            'Set rsFatFCE = Nothing
                            Exit Sub
                        End If
                        montaDadosVendas
                    End If
                    chamaForm.Show 1
                End If
            End If
        End If
        HabBotoes
        configControles
    End If
    Me.MousePointer = 0
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
    If apontaLV = 10 Then
        Dim i As Integer, leftPos As Single 'the left pos of the column
        Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate
        If Button = vbLeftButton Then
            If Not ListView1.SelectedItem Is Nothing Then
                ListView1.LabelEdit = lvwManual
                dx = GetLvwDeltaX
                lvwX = X + dx
                For i = 13 To 13
                    leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
                    If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                        m_RowIndex = ListView1.SelectedItem.Index 'row
                        m_ColIndex = i 'column
                            AlteraListview indiceVarGlobal
                            If varGlobal <> "" Then
                                
                                If ListView1.ListItems.Item(m_RowIndex).Checked And ListView1.SelectedItem.ListSubItems.Item(m_ColIndex).ReportIcon = "OK" Then
                                    mobjMsg.Abrir "Deseja ativar o registro para edição?", YesNo, pergunta, "Zeus"
                                    If Tp = 1 Then
                                        AtivaDesativaCago
                                    End If
                                End If
                                
                            End If
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
    Dim vStatus As Integer
    SqlAtivaDesativaCago = "update tbCD set status = case  WHEN status = 5 then 4 WHEN status = 4 then 5 ELSE 5 END where idcd = '" & Val(varGlobal) & "'"
    rsAtivaDesativaCago.Open SqlAtivaDesativaCago, cnBanco
    
    SqlAtivaDesativaCago = "Select status from tbCD where idcd = '" & Val(varGlobal) & "'"
    rsAtivaDesativaCago.Open SqlAtivaDesativaCago, cnBanco, adOpenKeyset, adLockReadOnly
    vStatus = rsAtivaDesativaCago.Fields(0)
    If vStatus <> 5 Then
        ListView1.SelectedItem.ListSubItems.Item(13) = ""
        ListView1.SelectedItem.ListSubItems.Item(13).ReportIcon = "AGUARDE-02"
    Else
        ListView1.SelectedItem.ListSubItems.Item(13) = ""
        ListView1.SelectedItem.ListSubItems.Item(13).ReportIcon = "OK"
    End If
    rsAtivaDesativaCago.Close
End Sub

'Private Sub achaDesenho()
'    Dim rsAchaDesenho As New ADODB.Recordset
'    Dim SqlAchaDesenho As String
'    Dim Y As Integer, X As Integer
'    Y = ListView1.ListItems.Count
'    For X = 1 To Y
'        ListView1.ListItems.Item(X).Selected = True
'        SqlAchaDesenho = "update tbmp set desenho = '" & ListView1.SelectedItem.ListSubItems.Item(7) & "' where idprogramacao = '" & Val(ListView1.ListItems.Item(X)) & "'"
'        rsAchaDesenho.Open SqlAchaDesenho, cnBanco
'    Next
'    Msgbox "Desenho inseridos com sucesso!"
'End Sub
