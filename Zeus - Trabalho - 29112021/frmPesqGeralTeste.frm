VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPesqGeralTeste 
   BorderStyle     =   0  'None
   ClientHeight    =   13905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20685
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13905
   ScaleWidth      =   20685
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   16960
      _Version        =   393216
      Style           =   1
      Tabs            =   20
      TabsPerRow      =   20
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPesqGeralTeste.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPesqGeralTeste.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPesqGeralTeste.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmPesqGeralTeste.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmPesqGeralTeste.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmPesqGeralTeste.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "frmPesqGeralTeste.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "frmPesqGeralTeste.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "frmPesqGeralTeste.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "frmPesqGeralTeste.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Tab 10"
      TabPicture(10)  =   "frmPesqGeralTeste.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Tab 11"
      TabPicture(11)  =   "frmPesqGeralTeste.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Tab 12"
      TabPicture(12)  =   "frmPesqGeralTeste.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "Tab 13"
      TabPicture(13)  =   "frmPesqGeralTeste.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).ControlCount=   0
      TabCaption(14)  =   "Tab 14"
      TabPicture(14)  =   "frmPesqGeralTeste.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).ControlCount=   0
      TabCaption(15)  =   "Tab 15"
      TabPicture(15)  =   "frmPesqGeralTeste.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).ControlCount=   0
      TabCaption(16)  =   "Tab 16"
      TabPicture(16)  =   "frmPesqGeralTeste.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).ControlCount=   0
      TabCaption(17)  =   "Tab 17"
      TabPicture(17)  =   "frmPesqGeralTeste.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).ControlCount=   0
      TabCaption(18)  =   "Tab 18"
      TabPicture(18)  =   "frmPesqGeralTeste.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).ControlCount=   0
      TabCaption(19)  =   "Tab 19"
      TabPicture(19)  =   "frmPesqGeralTeste.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).ControlCount=   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações "
      Height          =   8895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   16695
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa"
         Height          =   735
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   5175
         Begin VB.ComboBox Combo 
            Height          =   345
            ItemData        =   "frmPesqGeralTeste.frx":0230
            Left            =   120
            List            =   "frmPesqGeralTeste.frx":0232
            TabIndex        =   22
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox Text 
            Height          =   285
            Left            =   2400
            TabIndex        =   21
            Top             =   240
            Width           =   2055
         End
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
            MICON           =   "frmPesqGeralTeste.frx":0234
            PICN            =   "frmPesqGeralTeste.frx":0250
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.PictureBox picBg 
         Height          =   495
         Left            =   15600
         ScaleHeight     =   435
         ScaleMode       =   0  'User
         ScaleWidth      =   936.333
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   13
         Top             =   120
         Width           =   3975
         Begin ACTIVESKINLibCtl.SkinLabel Label3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmPesqGeralTeste.frx":09CA
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label4 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmPesqGeralTeste.frx":0A32
            TabIndex        =   15
            Top             =   480
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label2 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmPesqGeralTeste.frx":0A8C
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmPesqGeralTeste.frx":0AE6
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
      End
      Begin ZEUS.chameleonButton cmdconsulta0 
         Height          =   615
         Left            =   11040
         TabIndex        =   4
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
         MICON           =   "frmPesqGeralTeste.frx":0B4C
         PICN            =   "frmPesqGeralTeste.frx":0B68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta11 
         Height          =   615
         Left            =   9840
         TabIndex        =   5
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
         MICON           =   "frmPesqGeralTeste.frx":1842
         PICN            =   "frmPesqGeralTeste.frx":185E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta10 
         Height          =   615
         Left            =   9240
         TabIndex        =   6
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
         MICON           =   "frmPesqGeralTeste.frx":2538
         PICN            =   "frmPesqGeralTeste.frx":2554
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta8 
         Height          =   615
         Left            =   8640
         TabIndex        =   7
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
         MICON           =   "frmPesqGeralTeste.frx":322E
         PICN            =   "frmPesqGeralTeste.frx":324A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta9 
         Height          =   615
         Left            =   8040
         TabIndex        =   8
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
         MICON           =   "frmPesqGeralTeste.frx":3F24
         PICN            =   "frmPesqGeralTeste.frx":3F40
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta7 
         Height          =   615
         Left            =   1920
         TabIndex        =   9
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
         MICON           =   "frmPesqGeralTeste.frx":4C1A
         PICN            =   "frmPesqGeralTeste.frx":4C36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta6 
         Height          =   615
         Left            =   1320
         TabIndex        =   10
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
         MICON           =   "frmPesqGeralTeste.frx":5910
         PICN            =   "frmPesqGeralTeste.frx":592C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta5 
         Height          =   615
         Left            =   720
         TabIndex        =   11
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
         MICON           =   "frmPesqGeralTeste.frx":6606
         PICN            =   "frmPesqGeralTeste.frx":6622
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdconsulta4 
         Height          =   615
         Left            =   120
         TabIndex        =   12
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
         MICON           =   "frmPesqGeralTeste.frx":72FC
         PICN            =   "frmPesqGeralTeste.frx":7318
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
         TabIndex        =   23
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
      Begin ZEUS.chameleonButton cmdconsulta12 
         Height          =   615
         Left            =   10440
         TabIndex        =   24
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
         MICON           =   "frmPesqGeralTeste.frx":7FF2
         PICN            =   "frmPesqGeralTeste.frx":800E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   975
      Left            =   1680
      TabIndex        =   2
      Top             =   10080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   10080
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3960
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":8CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":99C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":A69C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":B376
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":C050
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":CD2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":DA04
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":E6DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":F3B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":10092
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":10D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":11A46
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":12720
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":133FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":140D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":14DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":15A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":16762
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1743C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":18116
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":18DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":19ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1A7A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1B47E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1C158
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1CE32
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1DB0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1E7E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":1F4C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":2019A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":20914
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":215EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":222C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":22FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":23C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":24956
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":25630
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":2630A
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":26FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":27CBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3360
      Top             =   9960
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
            Picture         =   "frmPesqGeralTeste.frx":28998
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":293AA
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":29DBC
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":2E7D6
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":331F0
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":41579
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":42253
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":42F2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":43C07
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":448E1
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":455BB
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":46295
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":46F6F
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":47C49
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":48923
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":495FD
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":4A2D7
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":4ACE9
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste.frx":4B9C3
            Key             =   "DUVIDA"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPesqGeralTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private WithEvents ctlDynamic As VBControlExtender
Private objText(19, 15) As TextBox
Attribute objText.VB_VarHelpID = -1
Private objFrame(19, 15) As Frame
Private objCombo(19, 15) As ComboBox
Private objLabel(19, 15) As Label
Private objListview(19, 15) As MSComctlLib.Listview
Attribute objListview.VB_VarHelpID = -1
Private objButton1(19, 15) As VBControlExtender
Attribute objButton1.VB_VarHelpID = -1
Private objButton(19, 15) As VBControlExtender
Private objPicture(19, 15) As PictureBox

Private WithEvents objTeste As VBControlExtender
Attribute objTeste.VB_VarHelpID = -1


Private Sub Command1_Click()
    Dim vProximaTab As Integer, X As Integer
    X = 19
    For vProximaTab = 0 To X
        If SSTab1.TabVisible(vProximaTab) = False Then
            Exit For
        Else
        End If
    Next
    If vProximaTab <= 19 Then
        SSTab1.TabVisible(vProximaTab) = True
        SSTab1.Tab = vProximaTab
        construirControles vProximaTab
    End If
End Sub

Private Sub Command2_Click()
    descontruirControles SSTab1.Tab
    SSTab1.TabVisible(SSTab1.Tab) = False
End Sub

Private Sub objTeste_Click()
    Msgbox "teste teste teste"
End Sub

Private Function construirControles(vTab As Integer)
    
    Set objFrame(vTab, 0) = Controls.Add("VB.Frame", "Frame1" + Trim(Str(vTab)), SSTab1)
    With objFrame(vTab, 0)
        .Visible = True
        .Top = 360
        .Left = 120
        .Width = 16695
        .Height = 9015
        .Caption = "Informações"
    End With

    Set objFrame(vTab, 1) = Controls.Add("VB.Frame", "Frame0" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objFrame(vTab, 1)
        .Visible = True
        .Top = 240
        .Left = 2760
        .Width = 5175
        .Height = 735
        .Caption = "Pesquisa"
    End With

    Set objCombo(vTab, 0) = Controls.Add("VB.ComboBox", "Combo" + Trim(Str(vTab)), objFrame(vTab, 1))
    With objCombo(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 120
        .Width = 2175
    End With

    Set objText(vTab, 0) = Controls.Add("VB.TextBox", "Text" + Trim(Str(vTab)), objFrame(vTab, 1))
    With objText(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 2400
        .Width = 2055
        .Height = 285
    End With

    Set objButton1(vTab, 0) = Controls.Add("zeus.chameleonButton", "chameleonButton0" + Trim(Str(vTab)), objFrame(vTab, 1))
    With objButton1(vTab, 0)
        .Visible = True
        .Top = 120
        .Left = 4560
        .Width = 495
        .Height = 495
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(30).Picture
        .Tag = "Pesquisar"
        .ToolTipText = "Pesquisar"
    End With

    Set objButton(vTab, 1) = Controls.Add("zeus.chameleonButton", "cmdconsulta4" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 1)
        .Visible = True
        .Top = 360
        .Left = 120
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(31).Picture
        .Tag = "Novo"
        .ToolTipText = "Novo"
    End With

    Set objButton(vTab, 2) = Controls.Add("zeus.chameleonButton", "cmdconsulta5" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 2)
        .Visible = True
        .Top = 360
        .Left = 720
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(32).Picture
        .Tag = "Editar"
        .ToolTipText = "Editar"
    End With

    Set objButton(vTab, 3) = Controls.Add("zeus.chameleonButton", "cmdconsulta6" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 3)
        .Visible = True
        .Top = 360
        .Left = 1320
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(33).Picture
        .Tag = "Excluir"
        .ToolTipText = "Excluir"
    End With

    Set objButton(vTab, 4) = Controls.Add("zeus.chameleonButton", "cmdconsulta7" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 4)
        .Visible = True
        .Top = 360
        .Left = 1920
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(34).Picture
        .Tag = "Sair"
        .ToolTipText = "Sair"
    End With

    Set objButton(vTab, 5) = Controls.Add("zeus.chameleonButton", "cmdconsulta9" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 5)
        .Visible = True
        .Top = 360
        .Left = 8040
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(35).Picture
        .Tag = "Admitir Candidato"
        .ToolTipText = "Admitir Candidato"
    End With

    Set objButton(vTab, 6) = Controls.Add("zeus.chameleonButton", "cmdconsulta8" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 6)
        .Visible = True
        .Top = 360
        .Left = 8640
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(36).Picture
        .Tag = "Filtro"
        .ToolTipText = "Filtro"
    End With

    Set objButton(vTab, 7) = Controls.Add("zeus.chameleonButton", "cmdconsulta10" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 7)
        .Visible = True
        .Top = 360
        .Left = 9240
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(37).Picture
        .Tag = "Imprimir"
        .ToolTipText = "Imprimir"
    End With

    Set objButton(vTab, 8) = Controls.Add("zeus.chameleonButton", "cmdconsulta11" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 8)
        .Visible = False
        .Top = 360
        .Left = 9840
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(38).Picture
        .Tag = "Atualiza Experiência"
        .ToolTipText = "Atualiza Experiência"
    End With

    Set objButton(vTab, 9) = Controls.Add("zeus.chameleonButton", "cmdconsulta12" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 9)
        .Visible = False
        .Top = 360
        .Left = 10440
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(39).Picture
        .Tag = "Afastamento/Retorno do Colaborador"
        .ToolTipText = "Afastamento/Retorno do Colaborador"
    End With

    Set objButton(vTab, 10) = Controls.Add("zeus.chameleonButton", "cmdconsulta0" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objButton(vTab, 10)
        .Visible = False
        .Top = 360
        .Left = 11040
        .Width = 615
        .Height = 615
        .Caption = ""
        .ButtonType = 11
        .PictureNormal = ImageList.ListImages(40).Picture
        .Tag = "Plano de Programação Semanal"
        .ToolTipText = "Afastamento/Retorno do Colaborador"
    End With

    Set objPicture(vTab, 0) = Controls.Add("VB.PictureBox", "picBg" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objPicture(vTab, 0)
        .Visible = False
        .Top = 360
        .Left = 15600
        .Width = 855
        .Height = 495
    End With
    
    
    Set objFrame(vTab, 2) = Controls.Add("VB.Frame", "Frame3" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objFrame(vTab, 2)
        .Visible = True
        .Top = 120
        .Left = 12360
        .Width = 3975
        .Height = 855
        .Caption = "Filtro "
        .Appearance = 0
        .BackColor = &H8000000F
    End With

    Set objLabel(vTab, 0) = Controls.Add("VB.Label", "Label1" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 0)
        .Visible = True
        .Top = 240
        .Left = 120
        .Width = 735
        .Height = 255
        .Caption = "Status: "
    End With

    Set objLabel(vTab, 1) = Controls.Add("VB.Label", "Label3" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 1)
        .Visible = True
        .Top = 480
        .Left = 120
        .Width = 855
        .Height = 255
        .Caption = "Período: "
    End With

    Set objLabel(vTab, 2) = Controls.Add("VB.Label", "Label2" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 2)
        .Visible = True
        .Top = 240
        .Left = 960
        .Width = 2055
        .Height = 255
        .Caption = "-"
    End With

    Set objLabel(vTab, 3) = Controls.Add("VB.Label", "Label4" + Trim(Str(vTab)), objFrame(vTab, 2))
    With objLabel(vTab, 3)
        .Visible = True
        .Top = 480
        .Left = 960
        .Width = 2055
        .Height = 255
        .Caption = "-"
    End With

    Set objListview(vTab, 0) = Controls.Add("MSComctlLib.ListViewCtrl.2", "Listview2" + Trim(Str(vTab)), objFrame(vTab, 0))
    With objListview(vTab, 0)
        .Visible = True
        .Top = 1080
        .Left = 120
        .Width = 16455
        .Height = 7695
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .LabelWrap = True
        .SortKey = 0
        .SortOrder = lvwAscending
        .View = lvwReport
        .BackColor = &H80000018
        .ForeColor = &H800000
    End With
    
    
'---------------
'    Set objTeste = Controls.Add("zeus.chameleonButton.1", "cmdTESTE", objFrame(vTab, 0))
'    With objTeste
'        .Visible = True
'        .Top = 360
'        .Left = 12000
'        .Width = 615
'        .Height = 615
'        .Caption = ""
'        .ButtonType = 11
'        .PictureNormal = ImageList.ListImages(38).Picture
'        .Tag = "teste"
'        .ToolTipText = "teste"
'    End With
    
    


End Function

Private Function descontruirControles(vTab As Integer)
On Error Resume Next
    Dim i As Long
    For i = 0 To 15
        Me.Controls.Remove objFrame(vTab, i).Name
        Me.Controls.Remove objText(vTab, i).Name
        Me.Controls.Remove objButton(vTab, i).Name
        Me.Controls.Remove objCombo(vTab, i).Name
        Me.Controls.Remove objLabel(vTab, i).Name
        Me.Controls.Remove objListview(vTab, i).Name
        Me.Controls.Remove objButton1(vTab, i).Name
        Me.Controls.Remove objPicture(vTab, i).Name
    Next
End Function

Private Function desconstroiTabs()
    Dim i As Long
    For i = 0 To 19
        SSTab1.TabVisible(i) = False
    Next
End Function


Private Sub Form_Load()


    desconstroiTabs
    'SSTab1.TabVisible(0) = True
End Sub





