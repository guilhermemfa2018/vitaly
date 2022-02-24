VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
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
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   0
         Left            =   11040
         TabIndex        =   20
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
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   12
         Left            =   10440
         TabIndex        =   19
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
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   11
         Left            =   9840
         TabIndex        =   18
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
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
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   10
         Left            =   9240
         TabIndex        =   17
         Tag             =   "Novo"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
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
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   8
         Left            =   8640
         TabIndex        =   16
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
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
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   9
         Left            =   8040
         TabIndex        =   15
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":4D98
         PICN            =   "frmPesqGeral.frx":4DB4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   7
         Left            =   1920
         TabIndex        =   13
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":5A8E
         PICN            =   "frmPesqGeral.frx":5AAA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   6
         Left            =   1320
         TabIndex        =   12
         Tag             =   "Cancelar"
         ToolTipText     =   "Cancelar"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":6784
         PICN            =   "frmPesqGeral.frx":67A0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   5
         Left            =   720
         TabIndex        =   14
         Tag             =   "Editar"
         ToolTipText     =   "Editar"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":747A
         PICN            =   "frmPesqGeral.frx":7496
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Tag             =   "Novo"
         ToolTipText     =   "Novo"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeral.frx":8170
         PICN            =   "frmPesqGeral.frx":818C
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
            NumListImages   =   22
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":8E66
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":9B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":A81A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":B4F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":C1CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":CEA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":DB82
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":E85C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":F536
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":10210
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":10EEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":11BC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":1289E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":13578
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":14252
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":14F2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":15C06
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":168E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":175BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":18294
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":18F6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":19C48
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
            NumListImages   =   22
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":376F3
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":38105
               Key             =   "EXC"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":38B17
               Key             =   "POSITIVO"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":3D531
               Key             =   "NEGATIVO"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":41F4B
               Key             =   "ARQUIVADO"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":502D4
               Key             =   "AGUARDE-01"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":50FAE
               Key             =   "AGUARDE-02"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":51C88
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":52962
               Key             =   "PENDENTE12"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5363C
               Key             =   "CONCLUIDO1"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":54316
               Key             =   "AVALIANDO1"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":54FF0
               Key             =   "PRETO"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":55CCA
               Key             =   "PENDENTE1"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":569A4
               Key             =   "FABRICANDO"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5767E
               Key             =   "FECHADO"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":58358
               Key             =   "EMPRESTAR"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":59032
               Key             =   "DEVOLVER"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":59D0C
               Key             =   "NAO"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5A9E6
               Key             =   "AST"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5BAC0
               Key             =   "EXC1"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5CB9A
               Key             =   "APR"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqGeral.frx":5DC74
               Key             =   "APP"
            EndProperty
         EndProperty
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
         Width           =   5500
         Begin ACTIVESKINLibCtl.SkinLabel Label3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmPesqGeral.frx":5ED4E
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label4 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmPesqGeral.frx":5EDB6
            TabIndex        =   9
            Top             =   480
            Width           =   2895
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label2 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmPesqGeral.frx":5EE10
            TabIndex        =   8
            Top             =   240
            Width           =   4800
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmPesqGeral.frx":5EE6A
            TabIndex        =   7
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
            ItemData        =   "frmPesqGeral.frx":5EED0
            Left            =   120
            List            =   "frmPesqGeral.frx":5EED2
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
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal fnBar As Long, lpScrollInfo As SCROLLINFO) As Long
 
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
Private X As Integer, Y As Integer
Private vCodCFO As String, vCodFilial As String, vDataCadMed As String
Private vDataCompetencia As String, vDataLimiteCompetencia As String, vNatFinanceira As String, vContaCaixa As String, vCondPagamento As String, vDataVencimento As String
Private vValor As String, vListaMedicoesReprovadas As String
Private vIDProduto As Integer, vStatusMedicaoEnv As String

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
    Dim Y As Integer, X As Integer
    Select Case Index
    Case 0
        
        If apontaLV = 9 Then
            'AlteraListview 1
            frmProgramacao.Show
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
        DesabBotoes
        varGlobal = 0
        Pesquisa = "novo"
        Status = "novo"
        If apontaLV = 0 Or apontaLV = 1 Then
            MarcaDesmarca ListView1
            If MeuLV.cmdconsulta(4).ToolTipText = "Desmarcar todos" Then
                Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(20).Picture
                MeuLV.cmdconsulta(4).ToolTipText = "Marcar todos"
            Else
                Set MeuLV.cmdconsulta(4).PictureNormal = MeuLV.ImageList1.ListImages(21).Picture
                MeuLV.cmdconsulta(4).ToolTipText = "Desmarcar todos"
            End If
            Exit Sub
        End If
        chamaForm.Show 1
        HabBotoes
        MontaLV (apontaLV)
    Case 5
        If apontaLV = 0 Then
            vQualColunaStatusMedicao = 10
        ElseIf apontaLV = 1 Then
            vQualColunaStatusMedicao = 8
        End If
        
        frmInformaDataExporta.Show 1
        If vDataExportMed <> "" Then
            'Msgbox vDataExportMed
            ExportaMedicao 1
        Else
            Msgbox "Procedimento cancelado.Nenhuma data de exportação informada para a medição"
        End If
    Case 6
        
        
'ConexaoLdap
        
        
        If apontaLV = 0 Then
            vQualColunaStatusMedicao = 10
        ElseIf apontaLV = 1 Then
            vQualColunaStatusMedicao = 8
        End If
        
        If apontaLV <> 9 And apontaLV <> 16 Then
            If AlteraListview(indiceVarGlobal) = False Then Exit Sub
        Else
            AlteraListview 2
        End If
        
        Pesquisa = "excluir"
        If MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) = "Aprovado" Or MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) = "Aprovado com Desconto" Then
            vDataExportMed = Date
            ExportaMedicao 2
        Else
            Msgbox "Aqui vai tratar medições ja exportadas"
            'ExcluirDadosLV apontaLV
        End If
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
        If apontaLV = 9 Or apontaLV = 12 Then cmdconsulta(9).Visible = True Else cmdconsulta(9).Visible = False
        HabBotoes
        Principal.StatusBar1.Panels(5).Text = "Registros: " & MeuLV.ListView1.ListItems.Count
    Case 9
        If apontaLV = 1 Then
            frmComposicaoNota.Show 1
            'AlteraListview 1
            'frmComunicacaoDesvio.Show 1
        ElseIf apontaLV = 5 Then
            frmRecFO.Show 1
        ElseIf apontaLV = 12 Then
            AlteraListview 1
            frmCausais.Show 1
        End If
    Case 10
        DesabBotoes
        Pesquisa = "Imprimir"
        If apontaLV = 9 Then
            frmPrintRels.Show 1
            'FCRConfronto.Show 1
        ElseIf apontaLV = 4 Then
            frmPrintRels.Show 1
            'FCRListaCargos.Show 1
        ElseIf apontaLV = 5 Then
            vDataBase = Date
            vDataCalc = CDate(vDataBase) - (vPeriodo * 30)
            criaTabTemp
            insereDadosTemp
            montaDadosClassifica
            AlteraListview 5
            frmPrintRels.Show 1
        ElseIf apontaLV = 6 Then
            vDataBase = Date
            vDataCalc = CDate(vDataBase) - (vPeriodo * 30)
            criaTabTemp
            insereDadosTemp
            montaDadosClassifica
            AlteraListview 1
            frmPrintRels.Show 1
        ElseIf apontaLV = 0 Then
            AlteraListview indiceVarGlobal
            frmPrintRels.Show 1
        ElseIf apontaLV = 18 Then
            'AlteraListview indiceVarGlobal
            'frmPrintRels.Show 1
        ElseIf apontaLV = 10 Then 'Programação
            'frmConvocacao.Show 1
        ElseIf apontaLV = 2 Or apontaLV = 3 Or apontaLV = 6 Or apontaLV = 11 Or apontaLV = 17 Then
            'FCRGeral.Show 1
        End If
        HabBotoes
    Case 11
        If apontaLV = 4 Then
            frmSelecionaGrupoAvFornec.Show 1
        ElseIf apontaLV = 9 Then
            AlteraListview 1
                frmFCE.Show 1
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
            If varGlobal2 <> "-" Then
                frmFCE.Show 1
            Else
                mobjMsg.Abrir "Nenhuma FCE selecionada", Ok, critico, "IMRM"
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
    If Err.Number = 400 Then
        Resume Next
    Else
        mobjMsg.Abrir "Nenhum item selecionado", Ok, critico, "Atenção"
    End If
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

Private Function AlteraListview(qtdCol As Integer)
    On Error GoTo Err
    AlteraListview = True
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    Dim vContaParaExclusao As Integer
    
    vContaParaExclusao = 0
    For X = 1 To Y
        
        If apontaLV = 13 Or apontaLV = 14 Then
           If ListView1.ListItems.Item(X).Selected = True Then
                'If ListView1.CheckBoxes = True Then ListView1.ListItems.Item(X).Checked = True
                Exit For
            End If
        Else
            MeuLV.ListView1.ListItems.Item(X).Selected = True
            If MeuLV.ListView1.ListItems.Item(X).Selected = True And MeuLV.ListView1.ListItems.Item(X).Checked = True Then
                vContaParaExclusao = vContaParaExclusao + 1
                Exit For
            End If
        End If
    Next
    
    
    
    
    
   If apontaLV <> 13 And apontaLV <> 14 Then
        'VERIFICA SE ALGUMA MEDIÇÃO FOI SELECIONADA PARA BLOQUEIO E RETORNA UMA MENSAGEM SE NÃO HOUVER SELEÇÃO
        If vContaParaExclusao = 0 Then
            mobjMsg.Abrir "Nenhuma medição foi selecionada para bloqueio!", Ok, informacao, "Atenção"
            AlteraListview = False
            Exit Function
        End If
        
        If MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) <> "Aprovado" And MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) <> "Aprovado com Desconto" And MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) <> "EXPORTADO" Then
            mobjMsg.Abrir "Medição selecionada não pode ser bloqueada. Status: " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao), Ok, informacao, "Atenção"
            AlteraListview = False
            Exit Function
        End If
   End If

    
    
    
    If Y > 0 Then
        If qtdCol = 1 Then
            varGlobal = ListView1.ListItems.Item(X)
        ElseIf qtdCol = 3 Then
            varGlobal = ListView1.SelectedItem.ListSubItems.Item(1)
        ElseIf qtdCol = 5 Then
            varGlobal = ListView1.SelectedItem.ListSubItems.Item(3)
        Else
            varGlobal = ListView1.ListItems.Item(X) & ListView1.SelectedItem.ListSubItems.Item(7)
        End If
        If apontaLV = 9 Then
            vRetrabalho = ListView1.SelectedItem.ListSubItems.Item(9)
        End If
        removeLinha = X
    End If
    Exit Function
Err:
    If Err.Number = 35600 Then
        Resume Next
    Else
        varGlobal = ""
        mobjMsg.Abrir "Nenhum registro cadastrado ou selecionado", Ok, critico, "Atenção"
        Exit Function
    End If
End Function

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
    If vEdi <> "N" Then
        'DesabBotoes
        Pesquisa = "editar"
        'If apontaLV = 16 Or apontaLV = 17 Then
        '    If apontaLV = 16 Then
        '        vSituacao = "INSPEÇÃO DE FABRICAÇÃO"
        '    Else
        '        vSituacao = "EXPEDIÇÃO"
        '    End If
        '    AlteraListview 2
        'Else
            AlteraListview indiceVarGlobal
        'End If
        
        If varGlobal <> "" Then
            'vTime = Time
            'vTime = RemoveMask(vTime)
            If apontaLV = 9 And vRetrabalho <> "-" Then
                frmRetrabalho.Show 1
            Else
                If apontaLV = 5 Or apontaLV = 13 Or apontaLV = 14 Or apontaLV = 16 Or apontaLV = 17 Then
                    'If apontaLV = 0 Then Set chamaForm = New frmEmprestimo
                    chamaForm.Show 1
                Else
                    If varGlobal2 = "?" Then
                        varGlobal2 = ""
                        Exit Sub
                    End If
                    If apontaLV = 5 Then
                        chamaForm.Show 1
                    End If
                    If apontaLV <> 0 And apontaLV <> 1 Then chamaForm.Show 1
                End If
            End If
        End If
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
                For i = 2 To 2
                    leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
                    If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                        m_RowIndex = ListView1.SelectedItem.Index 'row
                        m_ColIndex = i 'column
                            AlteraListview indiceVarGlobal
                            'If varGlobal <> "" Then AtivaDesativaCago
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
    GetScrollInfo ListView1.hwnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.hwnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Function ExportaMedicao(oQueFazer As Integer)
    Dim vContaMedicoesSelecionadas As Integer
    
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    ConexaoSAP
    vListaMedicoesReprovadas = ""
    For X = 1 To Y
        MeuLV.ListView1.ListItems.Item(X).Selected = True
        If MeuLV.ListView1.ListItems.Item(X).Selected = True And MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            'VERIFICA SE A MEDIÇÃO SELECIONADA ESTA REPROVADA
            If MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) <> "Aprovado" And MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) <> "Aprovado com Desconto" And MeuLV.ListView1.SelectedItem.ListSubItems.Item(vQualColunaStatusMedicao) <> "EXPORTADO" Then
                CompletaDados
                If vListaMedicoesReprovadas = "" Then
                    vListaMedicoesReprovadas = "(" & MeuLV.ListView1.ListItems.Item(X) & ")"
                Else
                    vListaMedicoesReprovadas = vListaMedicoesReprovadas & "; (" & MeuLV.ListView1.ListItems.Item(X) & ")"
                End If
            Else
                If oQueFazer = 1 Then
                    CompletaDados
                    If vStatusMedicaoEnv <> 1 Then
                        GeraIDMov
                        GeraNumeroMov MeuLV.ListView1.ListItems.Item(X)
                        If GravaMedicao(oQueFazer) = True Then
                            GravaTMov
                            gravaHistorico
                            GravaTitMMov
                            GravaRateioCC
                            AtualizaLV 1
                        End If
                    End If
                Else
                    CompletaDados
                    GravaMedicao oQueFazer
                    AtualizaLV 2
                End If
            End If
            
            vContaMedicoesSelecionadas = vContaMedicoesSelecionadas + 1
        End If
    Next
    
    cnBancoSAP.Close
    Set cnBancoSAP = Nothing
    
    If vContaMedicoesSelecionadas = 0 Then
        mobjMsg.Abrir "Nenhuma medição foi selecionada para exportação!", Ok, informacao, "Atenção"
        Exit Function
    Else
        'mobjMsg.Abrir vContaMedicoesSelecionadas & " medições selecionadas para exportação!", Ok, informacao, "Atenção"
    End If
    If vListaMedicoesReprovadas = "" Then
        If oQueFazer = 1 Then
            mobjMsg.Abrir "Todas as Medições foram exportadas com suscesso!", Ok, informacao, "Atenção"
        Else
            mobjMsg.Abrir "Cancelamento realizado!", Ok, exclamacao, "Atenção"
        End If
    Else
        Msgbox "As medições listadas abaixo não foram exportadas para o RM: " & vbCrLf & _
        vListaMedicoesReprovadas & "." & vbCrLf & vbCrLf & _
        "Pois as mesmas não constam como Aprovadas. " & vbCrLf & _
        "As demais foram aprovadas com sucesso. ", vbCritical, "Atenção"
        vListaMedicoesReprovadas = ""
    End If
End Function

Private Function GeraIDMov()
    Dim rsGeraIDMov As New ADODB.Recordset
    Dim SqlGeraIDMov As String
    
    Dim rsAtualizaIDMov As New ADODB.Recordset
    Dim SqlAtualizaIDMov As String
    
    SqlGeraIDMov = "select * from " & vBancoSAP & ".dbo.GAUTOINC as a where a.codautoinc like 'IDMOV' and a.codcoligada = '" & vCodColigadaRM & "'"
    rsGeraIDMov.Open SqlGeraIDMov, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraIDMov.RecordCount > 0 Then
        vIDMov = Val(rsGeraIDMov.Fields(3)) + 1
    Else
        vIDMov = 1
    End If
    rsGeraIDMov.Close
    Set rsGeraIDMov = Nothing
    
    Dim rsVerTmov As New ADODB.Recordset
    Dim SqlVerTmov As String
    Dim vOK As Boolean
    vOK = False
    While vOK = False
        SqlVerTmov = "select * from " & vBancoSAP & ".dbo.tMov where idmov = " & vIDMov
        rsVerTmov.Open SqlVerTmov, cnBanco, adOpenKeyset, adLockReadOnly
        If rsVerTmov.RecordCount > 0 Then
            vIDMov = vIDMov + 1
        Else
            vOK = True
        End If
        rsVerTmov.Close
    Wend
    Set rsGeraNumeroMov = Nothing
    
    SqlAtualizaIDMov = "UPDATE  " & vBancoSAP & ".dbo.GAUTOINC set VALAUTOINC = " & vIDMov & " where codautoinc like 'IDMOV' and codcoligada = '" & vCodColigadaRM & "'"
    rsAtualizaIDMov.Open SqlAtualizaIDMov, cnBanco
    Set rsAtualizaIDMov = Nothing
End Function

Private Function GeraNumeroMov(vNumMed As String)
    If apontaLV = 0 Then 'Significa que é medição de terceiros
        vNumeromov = RemoveMask2(Mid$(vNumMed, 10, 11), "-") 'Mid$(vNumMed, 1, 2) & Mid$(vNumMed, 10, 3)
        vNumeromov = RemoveMask2(vNumeromov, "/")
    Else
        vNumeromov = Format(vNumMed, "000000000")
    End If
End Function

Private Function GravaMedicao(statusMed As Integer)
On Error GoTo Err
    GravaMedicao = True
    
    Dim rsGravaMedicoes As New ADODB.Recordset
    Dim sqlGravaMedicoes As String
    Dim vTabMed As String
    
    If apontaLV = 0 Then
        vTabMed = "tbMedicoesTerceiro"
    Else
        vTabMed = "tbMedicoesPJ"
    End If
    sqlGravaMedicoes = "Select * from " & vTabMed
    rsGravaMedicoes.Open sqlGravaMedicoes, cnBanco, adOpenKeyset, adLockOptimistic
    
    
'    GravaTMov 'grava dados na tabela TMOV (TOTVS RM)
    
    rsGravaMedicoes.AddNew
    rsGravaMedicoes(1) = MeuLV.ListView1.ListItems.Item(X) 'Codigo da medição
    rsGravaMedicoes(2) = vIDMov 'Identificador do Movimento
    rsGravaMedicoes(4) = "-" 'Observação
    rsGravaMedicoes(5) = statusMed 'Stauts da medição (1 - Importada  2 - Rejeitada)
    rsGravaMedicoes(6) = vLogin 'Pega nome do login do usuário
    rsGravaMedicoes(7) = Format(vDataCadMed, "yyyy-mm-dd hh:mm:ss") 'Data de cadastro da medição
    rsGravaMedicoes(8) = vDataExportMed 'Data de exportação da Medição
        
'    GravaTitMMov X 'grava dados na tabela TITMMOV (TOTVS RM)
    If Not rsGravaMedicoes.EOF Then rsGravaMedicoes.Update
    rsGravaMedicoes.Close
    Exit Function
Err:
    If Err.Number = -2147217873 Then
        GravaMedicao = False
        mobjMsg.Abrir "A Medição que está tentando exportar já foi exportada", Ok, critico, "IMRM"
    End If
End Function

Private Sub GravaTMov()
    Dim rsGravaTMov As New ADODB.Recordset
    Dim SqlGravaTMov As String
   
    SqlGravaTMov = "Select A.CODCOLIGADA,A.IDMOV,A.CODFILIAL,A.CODLOC,A.CODCFO,A.CODCFONATUREZA,A.NUMEROMOV,A.SERIE,A.CODTMV,A.TIPO,A.STATUS,A.MOVIMPRESSO,A.DOCIMPRESSO,A.FATIMPRESSA,A.DATAEMISSAO,A.COMISSAOREPRES,A.VALORBRUTO,A.VALORLIQUIDO,A.VALOROUTROS,A.PERCCOMISSAO,A.PESOLIQUIDO," & _
    "A.PESOBRUTO,A.CODMOEVALORLIQUIDO,A.DATAMOVIMENTO,A.GEROUFATURA,A.CODCFOAUX,A.CODVEN1,A.CODVEN2,A.PERCCOMISSAOVEN2,A.CODCOLCFO,A.CODCOLCFONATUREZA,A.CODUSUARIO,A.GERADOPORLOTE,A.STATUSEXPORTCONT,A.GEROUCONTATRABALHO,A.GERADOPORCONTATRABALHO,A.HORULTIMAALTERACAO," & _
    "A.INDUSOOBJ,A.CONTABILIZADOPORTOTAL,A.INTEGRADOBONUM,A.FLAGPROCESSADO,A.ABATIMENTOICMS,A.USUARIOCRIACAO,A.DATACRIACAO,A.STSEMAIL,A.VALORBRUTOINTERNO,A.VINCULADOESTOQUEFL,A.VALORDESCCONDICIONAL,A.VALORDESPCONDICIONAL,A.CONTORCAMENTOANTIGO,A.SEQUENCIALESTOQUE," & _
    "A.INTEGRADOAUTOMACAO,A.INTEGRAAPLICACAO,A.DATALANCAMENTO,A.EXTENPORANEO,A.RECIBONFESTATUS,A.IDMOVCFO,A.VALORMERCADORIAS,A.USARATEIOVALORFIN,A.CODCOLCFOAUX,A.VRBASEINSSOUTRAEMPRESA,A.VALORBRUTOORIG,A.VALORLIQUIDOORIG,A.VALOROUTROSORIG,A.RECCREATEDBY,A.RECCREATEDON," & _
    "A.RECMODIFIEDBY,A.RECMODIFIEDON,A.DATASAIDA,A.DATAEXTRA1,A.CODTB1FLX,A.CODCPG,A.CODCXA,A.CODCOLCXA,A.CODFILIALDESTINO from tmov as a where a.idmov = '" & vIDMov & "'"
    rsGravaTMov.Open SqlGravaTMov, cnBancoSAP, adOpenKeyset, adLockOptimistic
 
    vValor = Format(vValor, "#,##0.00;(#,##0.00)")
    
    If rsGravaTMov.RecordCount = 0 Then
        rsGravaTMov.AddNew
        rsGravaTMov.Fields(0) = vCodColigadaRM 'Codigo da Coligada
        rsGravaTMov.Fields(1) = vIDMov 'Identificador
        rsGravaTMov.Fields(2) = vCodFilial 'Codigo da Filial
        rsGravaTMov.Fields(3) = Format(vCodFilial, "00")
        rsGravaTMov.Fields(4) = vCodCFO 'Codigo do Fornecedor
        rsGravaTMov.Fields(5) = "000001"
        rsGravaTMov.Fields(6) = vNumeromov 'NUMERO DO MOVIMENTO
        rsGravaTMov.Fields(7) = SerieEmpresa 'Numero de serie do movimento
        rsGravaTMov.Fields(8) = "1.1.06" 'Numero do movimento
        rsGravaTMov.Fields(9) = "A" 'Tipo do Movimento
        rsGravaTMov.Fields(10) = "R" 'Status do Movimento
        rsGravaTMov.Fields(11) = 0 'Movimento Impresso
        rsGravaTMov.Fields(12) = 0 'Documento Impresso
        rsGravaTMov.Fields(13) = 0 'Fatura Impressa
        rsGravaTMov.Fields(14) = Format(vDataCadMed, "yyyy-mm-dd hh:mm:ss") 'DATA DE EMISSAO - Data de cadastro da medição
        rsGravaTMov.Fields(15) = 0
        rsGravaTMov.Fields(16) = vValor 'Valor Bruto
        rsGravaTMov.Fields(17) = vValor 'Valor Liquido
        rsGravaTMov.Fields(18) = vValor 'Valor Outros
        rsGravaTMov.Fields(19) = 0
        rsGravaTMov.Fields(20) = 0
        rsGravaTMov.Fields(21) = 0
        rsGravaTMov.Fields(22) = "R$"
        rsGravaTMov.Fields(23) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss") 'Data do Movimento
        rsGravaTMov.Fields(24) = 0
        rsGravaTMov.Fields(25) = "CXXXXXXXXXX"
        'rsGravaTMov.Fields(26) = txtEmprestimo(0).Text
        'rsGravaTMov.Fields(27) = vCodVenRM
        rsGravaTMov.Fields(28) = 0
        rsGravaTMov.Fields(29) = 1
        'rsGravaTMov.Fields(30) = 1
        rsGravaTMov.Fields(31) = vLogin
        rsGravaTMov.Fields(32) = 0
        rsGravaTMov.Fields(33) = 0
        rsGravaTMov.Fields(34) = 0
        rsGravaTMov.Fields(35) = 0
        rsGravaTMov.Fields(36) = Time
        rsGravaTMov.Fields(37) = 0
        rsGravaTMov.Fields(38) = 0
        rsGravaTMov.Fields(39) = 0
        rsGravaTMov.Fields(40) = 0
        rsGravaTMov.Fields(41) = 0
        rsGravaTMov.Fields(42) = vLogin
        rsGravaTMov.Fields(43) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss") 'Data de criação do movimento
        rsGravaTMov.Fields(44) = 0
        rsGravaTMov.Fields(45) = vValor
        rsGravaTMov.Fields(46) = 0
        rsGravaTMov.Fields(47) = 0
        rsGravaTMov.Fields(48) = 0
        rsGravaTMov.Fields(49) = 0
        'rsGravaTMov.Fields(50) = vSequencialEstoque
        rsGravaTMov.Fields(51) = 0
        rsGravaTMov.Fields(52) = "T"
        rsGravaTMov.Fields(53) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss") 'Data da solicitação
        rsGravaTMov.Fields(54) = 0
        rsGravaTMov.Fields(55) = 0
        rsGravaTMov.Fields(56) = 539 '(verificar o que é)
        rsGravaTMov.Fields(57) = 0
        rsGravaTMov.Fields(58) = 0
        rsGravaTMov.Fields(59) = 1
        rsGravaTMov.Fields(60) = 0
        rsGravaTMov.Fields(61) = vValor
        rsGravaTMov.Fields(62) = vValor
        rsGravaTMov.Fields(63) = vValor
        rsGravaTMov.Fields(64) = vLogin
        rsGravaTMov.Fields(65) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss") '
        rsGravaTMov.Fields(66) = vLogin
        rsGravaTMov.Fields(67) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss") '
        rsGravaTMov.Fields(68) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss")
        rsGravaTMov.Fields(69) = Format(vDataCompetencia, "yyyy-mm-dd hh:mm:ss")
        If vNatFinanceira <> "" Then
            rsGravaTMov.Fields(70) = vNatFinanceira 'Natureza Financeira
        Else
            rsGravaTMov.Fields(70) = "3.08.0005" 'Natureza Financeira
        End If
        rsGravaTMov.Fields(71) = vCondPagamento 'Condição de pagamento
        rsGravaTMov.Fields(72) = vContaCaixa 'Conta/Caixa
        rsGravaTMov.Fields(73) = 1
        rsGravaTMov.Fields(74) = vCodFilial
    End If
    rsGravaTMov.Update
    rsGravaTMov.Close
    Set rsGravaTMov = Nothing
End Sub

Private Sub gravaHistorico()
    Dim rsGravaHistorico As New ADODB.Recordset
    Dim SqlGravaHistorico As String

    SqlGravaHistorico = "SELECT A.CODCOLIGADA,A.IDMOV,A.HISTORICOCURTO,A.RECCREATEDBY,A.RECCREATEDON,A.RECMODIFIEDBY,A.RECMODIFIEDON FROM TMOVHISTORICO AS A"
    rsGravaHistorico.Open SqlGravaHistorico, cnBancoSAP, adOpenKeyset, adLockOptimistic
    
    rsGravaHistorico.AddNew
    rsGravaHistorico.Fields(0) = vCodColigadaRM 'Codigo da Coligada
    rsGravaHistorico.Fields(1) = vIDMov 'Identificador
    If apontaLV = 0 Then
        rsGravaHistorico.Fields(2) = UCase("Ref. Servicos Prestados medicao " & MeuLV.ListView1.ListItems.Item(X) & " periodo " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) & " - " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(2)) 'Historico curto
    Else
        rsGravaHistorico.Fields(2) = UCase("Ref. Servicos Prestados medicao " & MeuLV.ListView1.ListItems.Item(X) & " periodo " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) & " - " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(2)) 'Historico curto
    End If
    rsGravaHistorico.Fields(3) = vLogin
    rsGravaHistorico.Fields(4) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss")
    rsGravaHistorico.Fields(5) = vLogin
    rsGravaHistorico.Fields(6) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss")
    
    rsGravaHistorico.Update
    rsGravaHistorico.Close
    Set rsGravaHistorico = Nothing
End Sub

Private Sub GravaTitMMov()
'On Error GoTo Err
    Dim rsGravaTitMMov As New ADODB.Recordset
    Dim SqlGravaTitMMov As String
   
    SqlGravaTitMMov = "SELECT A.CODCOLIGADA,A.IDMOV,A.NSEQITMMOV,A.NUMEROSEQUENCIAL,A.IDPRD,A.CODUND,A.QUANTIDADE,A.PRECOUNITARIO,A.VALORBRUTOITEM,A.RECCREATEDBY,A.RECCREATEDON,A.RECMODIFIEDBY,A.RECMODIFIEDON,A.QUANTIDADETOTAL,A.CODFILIAL,A.QUANTIDADEARECEBER,A.QUANTIDADEORIGINAL,A.QTDEVOLUMEUNITARIO FROM TITMMOV AS A WHERE A.IDMOV = '" & vIDMov & "'"
    rsGravaTitMMov.Open SqlGravaTitMMov, cnBancoSAP, adOpenKeyset, adLockOptimistic
    
    'If rsGravaTitMMov.RecordCount = 0 Then
        rsGravaTitMMov.AddNew
        rsGravaTitMMov.Fields(0) = vCodColigadaRM ' Código da coligada RM
        rsGravaTitMMov.Fields(1) = vIDMov 'identificador do movimento RM
        rsGravaTitMMov.Fields(2) = 1 ' Sequencial dos itens
        rsGravaTitMMov.Fields(3) = 1 ' Sequencial do itens
        rsGravaTitMMov.Fields(4) = vIDProduto 'Id do produto
        rsGravaTitMMov.Fields(5) = "UN" 'Unidade de Medida
        rsGravaTitMMov.Fields(6) = 1
        rsGravaTitMMov.Fields(7) = vValor
        rsGravaTitMMov.Fields(8) = vValor
        rsGravaTitMMov.Fields(9) = vLogin
        rsGravaTitMMov.Fields(10) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss")
        rsGravaTitMMov.Fields(11) = vLogin
        rsGravaTitMMov.Fields(12) = Format(vDataExportMed, "yyyy-mm-dd hh:mm:ss")
        rsGravaTitMMov.Fields(13) = 1 'QUANTIDADE TOTAL
        rsGravaTitMMov.Fields(14) = vCodFilial 'CODIGO DA FILIAL
        rsGravaTitMMov.Fields(15) = 1 'QUANTIDADE A RECEBER
        rsGravaTitMMov.Fields(16) = 1 'QUANTIDADE ORIGINAL
        rsGravaTitMMov.Fields(17) = 1 'QUANTIDADE VOLUME UNITARIO
        
        rsGravaTitMMov.Update
        rsGravaTitMMov.Close
        Set rsGravaTitMMov = Nothing
    'End If
    Exit Sub
Err:
    If Err.Number = -2147217873 Then
        mobjMsg.Abrir "O IDMOV nº" & vIDMov & " Já existe na tabela TMOV", Ok, critico, "IMRM"
        End
    End If
    Exit Sub
End Sub

Private Sub GravaRateioCC()
    Dim rsAchaRateioCC As New ADODB.Recordset
    Dim SqlAchaRateioCC As String
    
    Dim rsGravaRateioCC As New ADODB.Recordset
    Dim SqlGravaRateioCC As String
    
    Dim rsAtualizaIDMOVRATCCU As New ADODB.Recordset
    Dim SqlAtualizaIDMOVRATCCU As String
    
    Dim vIDMOVRATCCU As Long
    Dim vValorPercCC As String
    
    
    'PEGA O VALOR DO IDMOVRATCCU NA TABELA GAUTOINC
    SqlAtualizaIDMOVRATCCU = "select VALAUTOINC from GAUTOINC where codautoinc = 'IDMOVRATCCU' AND CODCOLIGADA = '" & vCodColigadaRM & "'"
    rsAtualizaIDMOVRATCCU.Open SqlAtualizaIDMOVRATCCU, cnBancoSAP, adOpenKeyset, adLockReadOnly
    vIDMOVRATCCU = rsAtualizaIDMOVRATCCU.Fields(0) + 1
    rsAtualizaIDMOVRATCCU.Close
    Set rsAtualizaIDMOVRATCCU = Nothing
    
    'TERCEIRO
    If apontaLV = 0 Then
        SqlAchaRateioCC = "select b.CODCCUSTO,b.PERCENTUALRATEIO from ID_APROP_MEDICAOTERCEIRO as a inner join ID_APROP_MEDICAOTERCEIRO_CC as b on a.id = b.IDMEDICAO where a.CODIGO = '" & MeuLV.ListView1.ListItems.Item(X) & "'"
        rsAchaRateioCC.Open SqlAchaRateioCC, cnBancoSAP, adOpenKeyset, adLockReadOnly
        
        
        While Not rsAchaRateioCC.EOF
            vValorPercCC = Format(vValor * rsAchaRateioCC.Fields(1) / 100, "#,##0.00;(#,##0.00)")
            vValorPercCC = Replace(vValorPercCC, ".", "")
            vValorPercCC = Replace(vValorPercCC, ",", ".")
            SqlGravaRateioCC = "Insert into TMOVRATCCU(IDMOV,CODCCUSTO,VALOR,IDMOVRATCCU,CODCOLIGADA) Values(" & vIDMov & ",'" & rsAchaRateioCC.Fields(0) & "'," & vValorPercCC & "," & vIDMOVRATCCU & "," & vCodcoligada & ")"
            rsGravaRateioCC.Open SqlGravaRateioCC, cnBancoSAP
            vIDMOVRATCCU = vIDMOVRATCCU + 1
            rsAchaRateioCC.MoveNext
        Wend
        vIDMOVRATCCU = vIDMOVRATCCU - 1
        rsAchaRateioCC.Close
        Set rsAchaRateioCC = Nothing
    'PJ/MENSAL
    Else
        Dim vDtIniPer As String, vDtTerPer As String
        vDtIniPer = Format(Mid$(MeuLV.ListView1.SelectedItem.ListSubItems.Item(7), 1, 10), "mm/dd/yyyy hh:mm:ss")
        vDtTerPer = Format(Mid$(MeuLV.ListView1.SelectedItem.ListSubItems.Item(7), 14, 10), "mm/dd/yyyy hh:mm:ss")
        SqlAchaRateioCC = "WITH MATRIZ_FILIAL AS  (SELECT CASE WHEN ID_APROPHORAS.IDPRJSGP IN (SELECT ID FROM ID_PRJ_PROJETO WHERE LEFT(RIGHT(CODIGO,12),4) IN ('9000')) THEN CONVERT(VARCHAR, CONVERT(INT, LEFT(ID_APROPHORAS.CODSECAO, 2))) WHEN (LEFT(ID_APROPHORAS.CODCCUSTO,6) = '1.0216' OR  LEFT(ID_APROPHORAS.CODCCUSTO,6) = '1.0217' ) THEN 3 ELSE CONVERT(VARCHAR, LEFT(ID_APROPHORAS.CODCCUSTO, 1)) END AS FILIAL," & _
                          "CASE WHEN CONDICAO = 'PJ' THEN QTDMINUTOS END AS HORAPJ,CASE WHEN CONDICAO = 'CLT' THEN QTDMINUTOS END AS HORACLT,QTDMINUTOS,CODCOLIGADA,DATAAPROPIADO,ID_APROPHORAS.CODCCUSTO FROM ID_APROPHORAS " & _
                          "WHERE  IDINFO = '" & MeuLV.ListView1.SelectedItem.ListSubItems.Item(17) & "' AND ID_APROPHORAS.DATAAPROPIADO BETWEEN '" & vDtIniPer & "' AND '" & vDtTerPer & "') " & _
                          "SELECT FILIAL,CODCCUSTO,SUM(QTDMINUTOS) QTDMINUTOS,convert(decimal(10,2),(convert(decimal(10,2),SUM(QTDMINUTOS) * 100) /  (SELECT SUM(QTDMINUTOS) FROM MATRIZ_FILIAL  WHERE  FILIAL = " & vCodFilial & "))) PERCENTUAL FROM MATRIZ_FILIAL " & _
                          "WHERE  FILIAL = " & vCodFilial & " GROUP BY FILIAL,CODCCUSTO "
        rsAchaRateioCC.Open SqlAchaRateioCC, cnBancoSAP, adOpenKeyset, adLockReadOnly
        While Not rsAchaRateioCC.EOF
            vValorPercCC = Format(vValor * rsAchaRateioCC.Fields(3) / 100, "#,##0.00;(#,##0.00)")
            vValorPercCC = Replace(vValorPercCC, ".", "")
            vValorPercCC = Replace(vValorPercCC, ",", ".")
            SqlGravaRateioCC = "Insert into TMOVRATCCU(IDMOV,CODCCUSTO,VALOR,IDMOVRATCCU,CODCOLIGADA) Values(" & vIDMov & ",'" & rsAchaRateioCC.Fields(1) & "'," & vValorPercCC & "," & vIDMOVRATCCU & "," & vCodcoligada & ")"
            rsGravaRateioCC.Open SqlGravaRateioCC, cnBancoSAP
            vIDMOVRATCCU = vIDMOVRATCCU + 1
            rsAchaRateioCC.MoveNext
        Wend
        vIDMOVRATCCU = vIDMOVRATCCU - 1
        rsAchaRateioCC.Close
        Set rsAchaRateioCC = Nothing
    End If
    'ATUALIZA O VALOR DO IDMOVRATCCU NA TABELA GAUTOINC
    SqlAtualizaIDMOVRATCCU = "UPDATE  " & vBancoSAP & ".dbo.GAUTOINC set VALAUTOINC = " & vIDMOVRATCCU & " where codautoinc like 'IDMOVRATCCU' and codcoligada = '" & vCodColigadaRM & "'"
    rsAtualizaIDMOVRATCCU.Open SqlAtualizaIDMOVRATCCU, cnBancoSAP
    Set rsAtualizaIDMOVRATCCU = Nothing
End Sub

Private Sub GravaFinanceiro()
    Dim rsGravaFinanceiro As New ADODB.Recordset
    Dim SqlGravaFinanceiro As String

    Dim rsGravaPercImposto As New ADODB.Recordset
    Dim SqlGravaPercImposto As String

    Dim rsAtualizaIDLANFin As New ADODB.Recordset
    Dim SqlAtualizaIDLANFin As String
    Dim vPercImpostos As String
    Dim vIDLAN As Long
    
    'PEGA O VALOR DO IDMOVRATCCU NA TABELA GAUTOINC
    SqlAtualizaIDLANFin = "select VALAUTOINC from GAUTOINC where codautoinc = 'IDLAN' AND CODCOLIGADA = '" & vCodColigadaRM & "'"
    rsAtualizaIDLANFin.Open SqlAtualizaIDLANFin, cnBancoSAP, adOpenKeyset, adLockReadOnly
    vIDLAN = rsAtualizaIDLANFin.Fields(0) + 1
    rsAtualizaIDLANFin.Close
    Set rsAtualizaIDLANFin = Nothing

    'INSERE OS DADOS NA TABELA FLAN
    
    SqlGravaFinanceiro = "SELECT A.CODFILIAL,A.CODCFO,A.CODTDO,A.NUMERODOCUMENTO,A.HISTORICO,A.SERIEDOCUMENTO,A.DATAEMISSAO,A.DATAVENCIMENTO,A.DATAPREVBAIXA,A.IDLAN,A.VALORORIGINAL,A.CODMOEVALORORIGINAL,A.CODCXA,A.IDCONVENIO,A.TIPOJUROSDIA,A.TIPOCONTABILLAN,A.CODTB1FLX,A.CODTB2FLX,A.DATAOP1,A.CODCOLIGADA,A.CODCOLCFO,A.CODCOLCXA,A.INSSEMOUTRAEMPRESA,A.REUTILIZACAO,A.IDMOV,A.PAGREC,A.STATUSLAN, " & _
                         "A.CODAPLICACAO,A.LIBAUTORIZADA,A.CNABACEITE,A.CNABBANCO,A.CATEGORIAAUTONOMO,A.VALORSERVIÇO,A.VRBASEINSS,A.VRBASEIRRF FROM FLAN AS A WHERE A.idlan = '" & vIDLAN & "'"
    rsGravaFinanceiro.Open SqlGravaFinanceiro, cnBancoSAP, adOpenKeyset, adLockOptimistic
    
    If rsGravaFinanceiro.RecordCount = 0 Then
        rsGravaFinanceiro.AddNew
        rsGravaFinanceiro.Fields(0) = vCodFilial 'CODIGO DA FILIAL
        rsGravaFinanceiro.Fields(1) = vCodCFO 'CODIGO DO FORNECEDOR
        rsGravaFinanceiro.Fields(2) = "034" 'CODIGO DO TIPO DO DOCUMENTO
        rsGravaFinanceiro.Fields(3) = vNumeromov & "/01" 'NUMERO DO DOCUMENTO
        If apontaLV = 0 Then
            rsGravaFinanceiro.Fields(4) = "MEDIÇÃO DE CONTRATO DE " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) 'HISTORICO COM O NOME DO FORNECEDOR
        Else
            rsGravaFinanceiro.Fields(4) = "MEDIÇÃO DE CONTRATO DE " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(18) 'HISTORICO COM O NOME DO FORNECEDOR
        End If
        rsGravaFinanceiro.Fields(5) = SerieEmpresa
        rsGravaFinanceiro.Fields(6) = Format(vDataCadMed, "yyyy-mm-dd hh:mm:ss") 'DATA DE EMISSAO
        rsGravaFinanceiro.Fields(7) = Format(vDataVencimento, "yyyy-mm-dd hh:mm:ss") 'DATA DE VENCIMENTO
        rsGravaFinanceiro.Fields(8) = Format(vDataVencimento, "yyyy-mm-dd hh:mm:ss") 'DATA DE PREVISÃO DE BAIXA
        rsGravaFinanceiro.Fields(9) = vIDLAN 'IDENTIFICADO DO LANÇAMENTO FINANCEIRO (IDLAN)
        rsGravaFinanceiro.Fields(10) = vValor 'VALOR ORIGINAL
        rsGravaFinanceiro.Fields(11) = "R$" 'MOEDA CORRENTE
        rsGravaFinanceiro.Fields(12) = vContaCaixa 'CONTA/CAIXA
        rsGravaFinanceiro.Fields(13) = 7 'IDENTIFICADOR DO CONVENIO
        rsGravaFinanceiro.Fields(14) = 0 'TIPO DE JUROS AO DIA
        rsGravaFinanceiro.Fields(15) = 2 'CONTABILIZAÇÃO - TIPO CONTABIL
        rsGravaFinanceiro.Fields(16) = vNatFinanceira 'NATUREZA FINANCEIRA
        rsGravaFinanceiro.Fields(17) = "008" 'FORMA DE PAGAMENTO
        rsGravaFinanceiro.Fields(18) = Format(vDataCompetencia, "yyyy-mm-dd hh:mm:ss") 'DATA DA COMPETENCIA
        rsGravaFinanceiro.Fields(19) = vCodcoligada 'CODIGO DA COLIGADA
        rsGravaFinanceiro.Fields(20) = vCodcoligada 'CODIGO DA COLIGADA
        rsGravaFinanceiro.Fields(21) = vCodcoligada 'CODIGO DA COLIGADA
        rsGravaFinanceiro.Fields(22) = 0
        rsGravaFinanceiro.Fields(23) = 0
        rsGravaFinanceiro.Fields(24) = vIDMov
        rsGravaFinanceiro.Fields(25) = 1 'PAGREC
        rsGravaFinanceiro.Fields(26) = 0 'STATUS DO LANÇAMENTO
        rsGravaFinanceiro.Fields(27) = "T" 'CODIGO DA APLICACAO
        rsGravaFinanceiro.Fields(28) = 0 'LIBAUTORIZADA
        rsGravaFinanceiro.Fields(29) = 0 'CNABACEITE
        rsGravaFinanceiro.Fields(30) = "341" 'CNABBANCO
        rsGravaFinanceiro.Fields(31) = 0 'CATEGORIAAUTONOMO
        rsGravaFinanceiro.Fields(32) = 0 'VALORSERVIÇO
        rsGravaFinanceiro.Fields(33) = 0 'VRBASEINSS
        rsGravaFinanceiro.Fields(34) = 0 'VRBASEIRRF
        
        rsGravaFinanceiro.Update
        rsGravaFinanceiro.Close
        Set rsGravaFinanceiro = Nothing
        
        vPercImpostos = "12.3100"
        SqlGravaPercImposto = "Insert into FLANCOMPL(CODCOLIGADA,IDLAN,PERC_IMPOSTO) Values(" & vCodcoligada & "," & vIDLAN & "," & vPercImpostos & ")"
        rsGravaPercImposto.Open SqlGravaPercImposto, cnBancoSAP
    End If
    
'    'ATUALIZA O VALOR DO IDMOVRATCCU NA TABELA GAUTOINC
    SqlAtualizaIDLANFin = "UPDATE  " & vBancoSAP & ".dbo.GAUTOINC set VALAUTOINC = " & vIDLAN & " where codautoinc like 'IDLAN' and codcoligada = '" & vCodColigadaRM & "'"
    rsAtualizaIDLANFin.Open SqlAtualizaIDLANFin, cnBancoSAP
    Set rsAtualizaIDLANFin = Nothing

End Sub

Private Sub CompletaDados()
    'Terceiro
    If apontaLV = 0 Then
        vCodCFO = MeuLV.ListView1.SelectedItem.ListSubItems.Item(13)
        vCodFilial = MeuLV.ListView1.SelectedItem.ListSubItems.Item(12)
        vDataCadMed = MeuLV.ListView1.SelectedItem.ListSubItems.Item(8)
        vValor = MeuLV.ListView1.SelectedItem.ListSubItems.Item(7)
        vDataCompetencia = "01/" & MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
        If MeuLV.ListView1.SelectedItem.ListSubItems.Item(14).ReportIcon = "OK" Then
            vStatusMedicaoEnv = 1
        Else
            vStatusMedicaoEnv = 0
        End If
    'PJ/Mensal
    Else
        vCodCFO = MeuLV.ListView1.SelectedItem.ListSubItems.Item(15)
        vCodFilial = MeuLV.ListView1.SelectedItem.ListSubItems.Item(14)
        vValor = MeuLV.ListView1.SelectedItem.ListSubItems.Item(12)
        'Paga data do aprovador
        If MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = "-" Then
            vDataCadMed = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
        Else
            vDataCadMed = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
        End If
        
        'DETERMINA COMPETENCIA DA MEDICAO DE PJ
        If DatePart("d", vDataCadMed) <= 10 Then
            vDataCompetencia = "01/" & Format(DatePart("m", Date) - 1, "00") & "/" & DatePart("yyyy", Date)
        End If
        If DatePart("d", vDataCadMed) > 10 Then
            If DatePart("m", vDataCadMed) < DatePart("m", Date) Then
                vDataCompetencia = "01/" & Format(DatePart("m", Date) - 1, "00") & "/" & DatePart("yyyy", Date)
            Else
                If DatePart("yyyy", vDataCadMed) < DatePart("yyyy", Date) Then
                    vDataCompetencia = "01/" & Format(DatePart("m", Date) - 1, "00") & "/" & DatePart("yyyy", Date)
                Else
                    vDataCompetencia = "01/" & Format(DatePart("m", Date), "00") & "/" & DatePart("yyyy", Date)
                End If
            End If
        End If
        vDataCompetencia = Format(vDataCompetencia, "dd/mm/yyyy")
        '------------------------------------
        
        If MeuLV.ListView1.SelectedItem.ListSubItems.Item(16).ReportIcon = "OK" Then
            vStatusMedicaoEnv = 1
        Else
            vStatusMedicaoEnv = 0
        End If
    End If
    'VERIFICA SE O FORNECEDOR É A IDG. SE FOR TRANSFERE PARA (VL SERVICOS)
    converteCFO
    
    
    'ACESSA A TABELA FCFODEF PARA PEGAR A NATUREZA FINANCEIRA E O NUMERO DA CONTA/CAIXA CADASTRADA PARA O FORNECEDOR
    Dim rsDadosFornecedor As New ADODB.Recordset
    Dim SqlDadosFornecedor As String

    SqlDadosFornecedor = "SELECT CODTB1FLX,CODCXA FROM FCFODEF WHERE CODCFO = '" & vCodCFO & "'"
    rsDadosFornecedor.Open SqlDadosFornecedor, cnBancoSAP, adOpenKeyset, adLockOptimistic
    If rsDadosFornecedor.RecordCount > 0 Then
        If Not IsNull(rsDadosFornecedor.Fields(0)) Then
            vNatFinanceira = rsDadosFornecedor.Fields(0)
        Else
            vNatFinanceira = "3.08.0005"
        End If
        If Not IsNull(rsDadosFornecedor.Fields(1)) Then
            vContaCaixa = rsDadosFornecedor.Fields(1)
        Else
            vContaCaixa = "341"
        End If
    Else
        vNatFinanceira = "3.08.0005"
        vContaCaixa = "341"
    End If
    rsDadosFornecedor.Close
    Set rsDadosFornecedor = Nothing
    
    'ACESSA A TABELA TCPGFCFO PARA PEGAR A CONDIÇÃO DE PAGAMENTO CADASTRADA PARA O FORNECEDOR
    SqlDadosFornecedor = "SELECT CODCPGCOMPRA FROM TCPGFCFO WHERE CODCFO = '" & vCodCFO & "'"
    rsDadosFornecedor.Open SqlDadosFornecedor, cnBancoSAP, adOpenKeyset, adLockOptimistic
    If rsDadosFornecedor.RecordCount > 0 Then
        If Not IsNull(rsDadosFornecedor.Fields(0)) Then
            vCondPagamento = rsDadosFornecedor.Fields(0)
        Else
            vCondPagamento = "80"
        End If
    Else
        vCondPagamento = "80"
    End If
    rsDadosFornecedor.Close
    Set rsDadosFornecedor = Nothing
   
    
    'ENCONTRA A CONDIÇÃO DE PAGAMENTO NA TABELA TCPG E REALIZA O CALCULO DA DATA DE VENCIMENTO
    Dim rsAchaCondPagamento As New ADODB.Recordset
    Dim SqlAchaCondPagamento As String
    
    SqlAchaCondPagamento = "select CODCPG, PRAZO1 AS PRAZO,TIPO1 as TIPO from TCPG where CODCPG = '" & vCondPagamento & "' and CODCOLIGADA = '" & vCodColigadaRM & "'"
    rsAchaCondPagamento.Open SqlAchaCondPagamento, cnBancoSAP, adOpenKeyset, adLockReadOnly
    If rsAchaCondPagamento.RecordCount > 0 Then
        Dim vDataInicioVencimento As String
        vDataInicioVencimento = Format(CDate(vDataExportMed) + 30, "yyyy-mm-dd hh:mm:ss")
        vDataInicioVencimento = "01/" & Format(DatePart("m", vDataInicioVencimento), "00") & "/" & DatePart("yyyy", vDataInicioVencimento)
        vDataVencimento = CDate(vDataInicioVencimento) + rsAchaCondPagamento.Fields(1)
    End If
    
    
    
    'ACESSA A TABELA TPRODUTO PARA PEGAR A O ID DO PRODUTO
    Dim rsDadosProduto As New ADODB.Recordset
    Dim SqlDadosProduto As String
    
    SqlDadosProduto = "select IDPRD from TPRODUTO where codigoprd = '01.01.0025' and CODCOLPRD = '" & vCodColigadaRM & "'"
    rsDadosProduto.Open SqlDadosProduto, cnBancoSAP, adOpenKeyset, adLockReadOnly
    If rsDadosProduto.RecordCount > 0 Then
        vIDProduto = rsDadosProduto.Fields(0)
    End If
    rsDadosProduto.Close
    Set rsDadosProduto = Nothing
End Sub

Private Sub converteCFO()
    If vCodCFO = "000229" Or vCodCFO = "001156" Or vCodCFO = "001724" Then
        vCodCFO = "002776"
    End If
End Sub

Private Sub AtualizaLV(escolha As Integer)
    If apontaLV = 0 Then
        If escolha = 1 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(14).ReportIcon = "OK"
            MeuLV.ListView1.ListItems.Item(X).Checked = False
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = "EXPORTADO"
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10).ForeColor = &HC0&
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(14).ReportIcon = "EXC"
            MeuLV.ListView1.ListItems.Item(X).Checked = False
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = "CANCELADO"
            MeuLV.ListView1.ListItems.Item(X).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(1).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(8).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(9).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(11).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(12).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(14).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(16).ForeColor = &H808080
        End If
    Else
        If escolha = 1 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(16).ReportIcon = "OK"
            MeuLV.ListView1.ListItems.Item(X).Checked = False
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = "EXPORTADO"
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(8).ForeColor = &HC0&
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(16).ReportIcon = "EXC"
            MeuLV.ListView1.ListItems.Item(X).Checked = False
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = "CANCELADO"
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(8).ForeColor = &HC0&
            MeuLV.ListView1.ListItems.Item(X).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(1).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(2).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(8).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(9).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(10).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(11).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(12).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(14).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(15).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(16).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(17).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(18).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(19).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(20).ForeColor = &H808080
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(21).ForeColor = &H808080
        End If
    End If
End Sub
