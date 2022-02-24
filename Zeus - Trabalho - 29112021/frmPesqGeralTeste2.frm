VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPesqGeralTeste2 
   BorderStyle     =   0  'None
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   28800
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
   ScaleHeight     =   10035
   ScaleWidth      =   28800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   8400
   End
   Begin VB.PictureBox picBg 
      Height          =   495
      Index           =   20
      Left            =   15000
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   8400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   8400
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   76050
      _ExtentX        =   134144
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   30
      TabsPerRow      =   30
      TabHeight       =   520
      TabMaxWidth     =   4410
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Carregando..."
      TabPicture(0)   =   "frmPesqGeralTeste2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView2(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClose(40)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Carregando..."
      TabPicture(1)   =   "frmPesqGeralTeste2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Carregando..."
      TabPicture(2)   =   "frmPesqGeralTeste2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Carregando..."
      TabPicture(3)   =   "frmPesqGeralTeste2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Carregando..."
      TabPicture(4)   =   "frmPesqGeralTeste2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Carregando..."
      TabPicture(5)   =   "frmPesqGeralTeste2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Carregando..."
      TabPicture(6)   =   "frmPesqGeralTeste2.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Carregando..."
      TabPicture(7)   =   "frmPesqGeralTeste2.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Carregando..."
      TabPicture(8)   =   "frmPesqGeralTeste2.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Carregando..."
      TabPicture(9)   =   "frmPesqGeralTeste2.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Carregando..."
      TabPicture(10)  =   "frmPesqGeralTeste2.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Carregando..."
      TabPicture(11)  =   "frmPesqGeralTeste2.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Carregando..."
      TabPicture(12)  =   "frmPesqGeralTeste2.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "Carregando..."
      TabPicture(13)  =   "frmPesqGeralTeste2.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).ControlCount=   0
      TabCaption(14)  =   "Carregando..."
      TabPicture(14)  =   "frmPesqGeralTeste2.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).ControlCount=   0
      TabCaption(15)  =   "Carregando..."
      TabPicture(15)  =   "frmPesqGeralTeste2.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).ControlCount=   0
      TabCaption(16)  =   "Carregando..."
      TabPicture(16)  =   "frmPesqGeralTeste2.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).ControlCount=   0
      TabCaption(17)  =   "Carregando..."
      TabPicture(17)  =   "frmPesqGeralTeste2.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).ControlCount=   0
      TabCaption(18)  =   "Carregando..."
      TabPicture(18)  =   "frmPesqGeralTeste2.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).ControlCount=   0
      TabCaption(19)  =   "Carregando..."
      TabPicture(19)  =   "frmPesqGeralTeste2.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).ControlCount=   0
      TabCaption(20)  =   "Carregando..."
      TabPicture(20)  =   "frmPesqGeralTeste2.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).ControlCount=   0
      TabCaption(21)  =   "Carregando..."
      TabPicture(21)  =   "frmPesqGeralTeste2.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).ControlCount=   0
      TabCaption(22)  =   "Carregando..."
      TabPicture(22)  =   "frmPesqGeralTeste2.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).ControlCount=   0
      TabCaption(23)  =   "Carregando..."
      TabPicture(23)  =   "frmPesqGeralTeste2.frx":0284
      Tab(23).ControlEnabled=   0   'False
      Tab(23).ControlCount=   0
      TabCaption(24)  =   "Carregando..."
      TabPicture(24)  =   "frmPesqGeralTeste2.frx":02A0
      Tab(24).ControlEnabled=   0   'False
      Tab(24).ControlCount=   0
      TabCaption(25)  =   "Carregando..."
      TabPicture(25)  =   "frmPesqGeralTeste2.frx":02BC
      Tab(25).ControlEnabled=   0   'False
      Tab(25).ControlCount=   0
      TabCaption(26)  =   "Carregando..."
      TabPicture(26)  =   "frmPesqGeralTeste2.frx":02D8
      Tab(26).ControlEnabled=   0   'False
      Tab(26).ControlCount=   0
      TabCaption(27)  =   "Carregando..."
      TabPicture(27)  =   "frmPesqGeralTeste2.frx":02F4
      Tab(27).ControlEnabled=   0   'False
      Tab(27).ControlCount=   0
      TabCaption(28)  =   "Carregando..."
      TabPicture(28)  =   "frmPesqGeralTeste2.frx":0310
      Tab(28).ControlEnabled=   0   'False
      Tab(28).ControlCount=   0
      TabCaption(29)  =   "Carregando..."
      TabPicture(29)  =   "frmPesqGeralTeste2.frx":032C
      Tab(29).ControlEnabled=   0   'False
      Tab(29).ControlCount=   0
      Begin ZEUS.chameleonButton cmdClose 
         Height          =   255
         Index           =   40
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   11
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPesqGeralTeste2.frx":0348
         PICN            =   "frmPesqGeralTeste2.frx":0364
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6135
         Index           =   20
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   16575
         _ExtentX        =   29236
         _ExtentY        =   10821
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4080
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   41
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":0780
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2134
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":47C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":549C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6176
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8804
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":94DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A1B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AE92
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":BB6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":C846
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":D520
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":E1FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":EED4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":FBAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":10888
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":11562
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1223C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":12F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":13BF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":148CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":155A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1627E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":16F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":17C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":183AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":19086
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":19D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1AA3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1B714
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1C3EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1D0C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1DDA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1EA7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F756
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1FF98
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3240
      Top             =   9000
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
            Picture         =   "frmPesqGeralTeste2.frx":21072
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":21A84
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":22496
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":26EB0
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2B8CA
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":39C53
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3A92D
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3B607
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3C2E1
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3CFBB
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3DC95
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3E96F
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3F649
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":40323
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":40FFD
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":41CD7
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":429B1
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":433C3
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4409D
            Key             =   "DUVIDA"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmPesqGeralTeste2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private vSSTab As SSTab

Private Function descontruirControles(vTab As Integer)
On Error Resume Next
    Dim i As Long
    For i = 0 To 15
        Me.Controls.Remove objFrame(vTab, i).Name
        Me.Controls.Remove objText(vTab, i).Name
        Me.Controls.Remove objButton1(vTab, i).Name
        Me.Controls.Remove objCombo(vTab, i).Name
        Me.Controls.Remove objLabel(vTab, i).Name
        Me.Controls.Remove objListview(vTab, i).Name
        Me.Controls.Remove objPicture(vTab, i).Name
    Next
End Function

Private Sub cmdClose_Click(Index As Integer)
    If vListViewPrincipal.ListItems.Count > 0 Then GravarConfLVTeste
    Principal.StatusBar1.Panels(5).Text = "Registros: "
    desconstruirBotao SSTab1.Tab
    If contaTabsAbertas(SSTab1) = 1 Then
        'vLeftPadrao = vLeftPadrao - 2520
        tabAberta = False
        Unload Me
    Else
        SSTab1.TabVisible(SSTab1.Tab) = False
        'construirBotaoClose frmPesqGeralTeste2.SSTab1
    End If
End Sub

Private Sub Form_Load()
    vPosAtual = 1
    Set mo_Events = New Collection
    Set vSSTab = SSTab1
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    desconstroiTabs vSSTab
    SubClassSSTAB SSTab1, Picture1
'    acertaTamanhoIcone
'    DimensionaLV "Métodos e Processos"
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Public Sub ImgClick(p_idx As Long)
    setaComponentesTab SSTab1
    vTime = Time
    vTime = RemoveMask(vTime)
    Select Case p_idx
'-- NOVO
        Case 1, 11, 21, 31, 41, 51, 61, 71, 81, 91, 101, 111, 121, 131, 141, 151, 161, 171, 181, 191
            If apontaLV = 16 Then
                AlteraListview 2, vListViewPrincipal
            Else
                AlteraListview 1, vListViewPrincipal
            End If
            If varGlobal2 = "?" Then
                varGlobal2 = ""
                Exit Sub
            End If
            
            Pesquisa = "novo"
            Status = "novo"
            If apontaLV = 6 Then
                AlteraListview indiceVarGlobal, vListViewPrincipal
                frmLM.Show 1
            Else
                If apontaLV = 16 Or apontaLV = 17 Then
                    If apontaLV = 16 Then
                        vSituacao = "INSPEÇÃO DE FABRICAÇÃO"
                    Else
                        vSituacao = "EXPEDIÇÃO"
                        Set chamaForm = New frmRelExp
                    End If
                    AlteraListview 2, vListViewPrincipal
                    chamaForm.Show
                    Exit Sub
                Else
                    chamaForm.Show 1
                End If
            End If
            
'-- EDITAR
        Case 2, 12, 22, 32, 42, 52, 62, 72, 82, 92, 102, 112, 122, 132, 142, 152, 162, 172, 182, 192
            AlteraListview 1, vListViewPrincipal
            If varGlobal2 = "?" Then
                mobjMsg.Abrir "FCE concluida. A LM não pode ser EDITADA", Ok, critico, "Atenção"
                varGlobal2 = ""
                Exit Sub
            End If

            If apontaLV = 18 Then
                Unload Me
                Exit Sub
            End If
            Pesquisa = "editar"
            AlteraListview indiceVarGlobal, vListViewPrincipal
            If varGlobal <> "" Then
                If apontaLV = 9 And vRetrabalho <> "-" Then
                    frmRetrabalho.Show 1
                ElseIf apontaLV = 17 Then
                    vSituacao = "EXPEDIÇÃO TERC."
                    frmRelExpAvulso.Show
                    Exit Sub
                Else
                    chamaForm.Show 1
                    Exit Sub
                End If
            End If

'-- EXCLUIR
        Case 3, 13, 23, 33, 43, 53, 63, 73, 83, 93, 103, 113, 123, 133, 143, 153, 163, 173, 183, 193
            If apontaLV <> 9 And apontaLV <> 16 Then
                AlteraListview indiceVarGlobal, vListViewPrincipal
            Else
                AlteraListview 2, vListViewPrincipal
            End If
            If apontaLV = 16 Then
                vSituacao = "INSPEÇÃO DE PINTURA"
                chamaForm.Show
                Exit Sub
            End If
            
            Pesquisa = "excluir"
            'SERÁ REALIZADA A ADAPTAÇÃO DA FUNCTION ABAIXO CHAMADA... PARA QUE POSSA SER EXECUTADA DINAMICAMENTE
            'CarregaSQLExcluir apontaLV
            If apontaLV <> 11 And apontaLV <> 6 And apontaLV <> 5 And apontaLV <> 4 And apontaLV <> 3 And apontaLV <> 2 And apontaLV <> 0 And apontaLV <> 10 And apontaLV <> 9 And apontaLV <> 8 And apontaLV <> 15 And apontaLV <> 18 And apontaLV <> 19 Then ExcluirDadosLV
        
'-- SAIR
        Case 4, 14, 24, 34, 44, 54, 64, 74, 84, 94, 104, 114, 124, 134, 144, 154, 164, 174, 184, 194
            If vListViewPrincipal.ListItems.Count > 0 Then GravarConfLVTeste
            Principal.StatusBar1.Panels(5).Text = "Registros: "
            desconstruirBotao SSTab1.Tab
            If contaTabsAbertas(SSTab1) = False And tabAberta = False Then
                tabAberta = False
                Unload Me
            ElseIf contaTabsAbertas(SSTab1) = False And tabAberta = True Then
                SSTab1.TabVisible(SSTab1.Tab) = False
            ElseIf contaTabsAbertas(SSTab1) = True And tabAberta = True Then
                SSTab1.TabVisible(SSTab1.Tab) = False
                'construirBotaoClose frmPesqGeralTeste2.SSTab1
            End If
'-- CD - COMUNICAÇÃO DE DESVIO | RECEBER FO | CAUSAIS   | ADMITIR CANDIDATO
        Case 5, 15, 25, 35, 45, 55, 65, 75, 85, 95, 105, 115, 125, 135, 145, 155, 165, 175, 185, 195
            If apontaLV = 9 Then
                frmComunicacaoDesvio.Show 1
            ElseIf apontaLV = 5 Then
                frmRecFO.Show 1
            ElseIf apontaLV = 12 Then
                AlteraListview 1, vListViewPrincipal
                frmCausais.Show 1
            End If
'-- FILTRAR
        Case 6, 16, 26, 36, 46, 56, 66, 76, 86, 96, 106, 116, 126, 136, 146, 156, 166, 176, 186, 196
            FiltroGeral = ""
            Tipo = False
            Pesquisa = "filtro"
            frmFiltro.Combo1 = "TODOS"
            

            
            quantidadeDeFrom apontaLV
            
            frmFiltro.Show 1
            If apontaLV = 0 Then
                'vQdtFrom = 1
                MontaCabLV "Código", "Descrição", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListViewTipoMaterial
                MontaDadosLVTeste "S", vListViewTipoMaterial
                PersonaColLVTeste 2, "N", "N", "", "S", "N", "N", "E", vListViewTipoMaterial
                If vListViewTipoMaterial.ListItems.Count > 0 Then ajusta_LVTeste vListViewTipoMaterial
            End If
            If apontaLV = 1 Then
                'vQdtFrom = 1
                MontaCabLV "Código", "Nome", "Endereço", "CEP", "Bairro", "Cidade", "UF", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewClientes
                MontaDadosLVTeste "S", vListviewClientes
                PersonaColLVTeste 7, "N", "P", "", "S", "N", "N", "E", vListviewClientes
                If vListviewClientes.ListItems.Count > 0 Then ajusta_LVTeste vListviewClientes
            End If
            If apontaLV = 2 Then
                'vQdtFrom = 1
                MontaCabLV "ID", "Tipo", "Código", "Nome", "Descrição", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListViewParadas
                MontaDadosLVTeste "S", vListViewParadas
                PersonaColLVTeste 5, "N", "P", "", "S", "N", "N", "E", vListViewParadas
                If vListViewParadas.ListItems.Count > 0 Then ajusta_LVTeste vListViewParadas
            End If
            If apontaLV = 3 Then
                'vQdtFrom = 1
                MontaCabLV "Código", "Nome", "CNPJ", "IE", "Endereço", "CEP", "Bairro", "Cidade", "UF", "Ativo", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewTransportadoras
                MontaDadosLVTeste "N", vListviewTransportadoras
                PersonaColLVTeste 9, "N", "P", "", "S", "N", "N", "E", vListviewTransportadoras
                If vListviewTransportadoras.ListItems.Count > 0 Then ajusta_LVTeste vListviewTransportadoras
            End If
            If apontaLV = 4 Then
                'vQdtFrom = 1
                MontaCabLV "ID", "Código", "Descrição", "Cod Tipo", "Tipo Material", "Fórmula PESO", "Fórmula PINTURA", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewFormulaPRD
                MontaDadosLVTeste "S", vListviewFormulaPRD
                If vListviewFormulaPRD.ListItems.Count > 0 Then ajusta_LVTeste vListviewFormulaPRD
            End If
            If apontaLV = 5 Then
                'vQdtFrom = 4
                MontaCabLV "FO", "Empresa", "Coleta nº", "Contato", "Fone", "Descrição", "Data Abertura", "Dev. CP", "Proposta nº", "Quant.", "Valor Unit", "Valor Total", "Pedido nº", "FCE nº", "Status FO", "Ativo", "Status FCE", "Tipo FCE", "", "", ""
                contaColLVTeste vListviewComercial
                MontaDadosLVTeste "S", vListviewComercial
                PersonaColLVTeste 14, "N", "N", "", "S", "N", "N", "E", vListviewComercial
                PersonaColLVTeste 13, "S", "S", "", "N", "N", "N", "D", vListviewComercial
                PersonaColLVTeste 15, "N", "N", "", "S", "N", "N", "E", vListviewComercial
                PersonaColLVTeste 16, "N", "N", "", "S", "N", "N", "E", vListviewComercial
                PersonaColLVTeste 17, "N", "P", "", "N", "N", "N", "E", vListviewComercial 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
                If vListviewComercial.ListItems.Count > 0 Then ajusta_LVTeste vListviewComercial
            End If
            If apontaLV = 6 Then
                'vQdtFrom = 3
                MontaCabLV "Data abertura", "FCE", "Cliente", "Contato", "Fone", "Data entrega", "Pintura", "Transporte", "Matéria-prima", "Fabricação", "Reparo", "Ativo", "Status FCE", "Tipo FCE", "", "", "", "", "", "", ""
                contaColLVTeste vListviewFCE
                MontaDadosLVTeste "S", vListviewFCE
                PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "D", vListviewFCE
                PersonaColLVTeste 11, "N", "N", "", "S", "N", "N", "E", vListviewFCE
                PersonaColLVTeste 12, "N", "N", "", "S", "N", "N", "E", vListviewFCE
                PersonaColLVTeste 13, "N", "P", "", "N", "N", "N", "E", vListviewFCE 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
                If vListviewFCE.ListItems.Count > 0 Then ajusta_LVTeste vListviewFCE
            End If
            If apontaLV = 7 Then
                'vQdtFrom = 3
                MontaCabLV "Identificador", "Desenho", "Rev.", "FCE", "Projeto", "Data Cadastro", "Tipo", "Ativo", "Tipo FCE", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewDesenhos
                MontaDadosLVTeste "S", vListviewDesenhos
                PersonaColLVTeste 1, "S", "N", "", "N", "N", "N", "E", vListviewDesenhos
                PersonaColLVTeste 7, "N", "N", "", "S", "N", "N", "E", vListviewDesenhos
                PersonaColLVTeste 8, "N", "P", "", "N", "N", "N", "E", vListviewDesenhos 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
                If vListviewDesenhos.ListItems.Count > 0 Then ajusta_LVTeste vListviewDesenhos
            End If
            If apontaLV = 8 Then
                'vQdtFrom = 3
                MontaCabLV "FCE", "LM", "Data Abertura", "Descrição", "Ativo", "Status FCE", "Tipo FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewLM
                MontaDadosLVTeste "S", vListviewLM
                PersonaColLVTeste 1, "S", "S", "", "N", "S", "N", "D", vListviewLM
                PersonaColLVTeste 4, "N", "N", "", "S", "N", "N", "E", vListviewLM
                PersonaColLVTeste 5, "N", "N", "", "S", "N", "N", "E", vListviewLM
                PersonaColLVTeste 6, "N", "P", "", "N", "N", "N", "E", vListviewLM  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
                If vListviewLM.ListItems.Count > 0 Then ajusta_LVTeste vListviewLM
            End If
            If apontaLV = 9 Then
                'vQdtFrom = 3
                MontaCabLV "Planejamento", "OS nº", "Rev.", "Data", "FCE", "Projeto", "Responsável", "Desenho", "Ativo", "Retrabalho", "Status", "Status FCE", "Tipo FCE", "Tipo OS", "", "", "", "", "", "", ""
                contaColLVTeste vListviewMP
                MontaDadosLVTeste "S", vListviewMP
                PersonaColLVTeste 1, "N", "N", "", "N", "S", "N", "E", vListviewMP
                PersonaColLVTeste 8, "N", "N", "", "S", "N", "N", "E", vListviewMP
                PersonaColLVTeste 9, "S", "N", "", "N", "N", "N", "E", vListviewMP
                PersonaColLVTeste 10, "N", "S", "", "N", "N", "N", "E", vListviewMP
                PersonaColLVTeste 11, "N", "N", "", "S", "N", "N", "E", vListviewMP
                PersonaColLVTeste 12, "N", "P", "", "N", "N", "N", "E", vListviewMP  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
                If vListviewMP.ListItems.Count > 0 Then ajusta_LVTeste vListviewMP
            End If
            If apontaLV = 10 Then
                'vQdtFrom = 1
                MontaCabLV "Identificador", "FCE", "Desenho", "Rev.", "Quant.", "Peso Unit.", "Peso Total", "Recebido", "Previsão Det.", "Usuário", "Data inicio", "Data fim", "Croqui", "Status", "Observação", "Ativo", "Detalhista", "", "", "", ""
                contaColLVTeste vListviewControleDesenhos
                MontaDadosLVTeste "S", vListviewControleDesenhos
                PersonaColLVTeste 1, "S", "N", "", "N", "N", "N", "E", vListviewControleDesenhos
                PersonaColLVTeste 4, "N", "N", "", "N", "N", "N", "D", vListviewControleDesenhos
                PersonaColLVTeste 5, "N", "N", "", "N", "N", "S", "D", vListviewControleDesenhos
                PersonaColLVTeste 6, "N", "N", "", "N", "N", "S", "D", vListviewControleDesenhos
                PersonaColLVTeste 13, "N", "N", "", "S", "N", "N", "E", vListviewControleDesenhos
                PersonaColLVTeste 15, "N", "P", "", "S", "N", "N", "E", vListviewControleDesenhos
                If vListviewControleDesenhos.ListItems.Count > 0 Then ajusta_LVTeste vListviewControleDesenhos
            End If
            If apontaLV = 11 Then
                'vQdtFrom = 1
                MontaCabLV "Centro de Custo", "Nome Centro de Custo", "Fórmula", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewFormulaCC
                MontaDadosLVTeste "S", vListviewFormulaCC
                If vListviewFormulaCC.ListItems.Count > 0 Then ajusta_LVTeste vListviewFormulaCC
            End If
            If apontaLV = 12 Then
                'vQdtFrom = 1
                MontaCabLV "CD nº", "Data Abertura", "Responsável", "OS nº", "FCE", "Projeto", "Observação", "Status", "RNC nº", "Data Conclusão", "Retrabalho", "Retrabalho nº", "Data Fechamento", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewRNCF
                MontaDadosLVTeste "S", vListviewRNCF
                PersonaColLVTeste 3, "N", "N", "", "N", "S", "N", "E", vListviewRNCF
                PersonaColLVTeste 7, "S", "S", "", "S", "N", "N", "E", vListviewRNCF
                PersonaColLVTeste 8, "S", "N", "", "N", "S", "N", "E", vListviewRNCF
                PersonaColLVTeste 10, "N", "P", "", "S", "N", "N", "E", vListviewRNCF
                PersonaColLVTeste 11, "S", "S", "", "N", "S", "N", "E", vListviewRNCF
                If vListviewRNCF.ListItems.Count > 0 Then ajusta_LVTeste vListviewRNCF
            End If
            If apontaLV = 13 Then
                'vQdtFrom = 1
                MontaCabLV "Código", "Nome do usuário", "Grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewUsuarios
                MontaDadosLVTeste "S", vListviewUsuarios
                PersonaColLVTeste 3, "N", "P", "", "S", "N", "N", "E", vListviewUsuarios
                If vListviewUsuarios.ListItems.Count > 0 Then ajusta_LVTeste vListviewUsuarios
            End If
            If apontaLV = 14 Then
                'vQdtFrom = 1
                MontaCabLV "Código", "Nome do grupo", "Ativo", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewGrupos
                MontaDadosLVTeste "S", vListviewGrupos
                PersonaColLVTeste 2, "N", "P", "", "S", "N", "N", "E", vListviewGrupos
                If vListviewGrupos.ListItems.Count > 0 Then ajusta_LVTeste vListviewGrupos
            End If
            If apontaLV = 15 Then
                'vQdtFrom = 3
                MontaCabLV "Chapa", "Nome", "Permissão", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewPermissoes
                MontaDadosLVTeste "S", vListviewPermissoes
                PersonaColLVTeste 2, "N", "P", "", "S", "N", "N", "E", vListviewPermissoes
                If vListviewPermissoes.ListItems.Count > 0 Then ajusta_LVTeste vListviewPermissoes
            End If
            If apontaLV = 16 Then
                'vQdtFrom = 1
                MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "Nome Cliente", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewRelInsp
                MontaDadosLVTeste "S", vListviewRelInsp
                PersonaColLVTeste 4, "N", "P", "", "S", "N", "N", "E", vListviewRelInsp
                If vListviewRelInsp.ListItems.Count > 0 Then ajusta_LVTeste vListviewRelInsp
            End If
            If apontaLV = 17 Then
                'vQdtFrom = 1
                MontaCabLV "ID Proj.", "FCE", "Projeto (TAG/Pacote/Elemento)", "Nome Cliente", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewRelExpedicao
                MontaDadosLVTeste "S", vListviewRelExpedicao
                PersonaColLVTeste 4, "N", "P", "", "S", "N", "N", "E", vListviewRelExpedicao
                If vListviewRelExpedicao.ListItems.Count > 0 Then ajusta_LVTeste vListviewRelExpedicao
            End If
            If apontaLV = 18 Then
                'vQdtFrom = 1
                MontaCabLV "Nº Relatório", "FCE", "Projeto", "Descrição", "Data emissão", "Status Impressão", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewImpExpedicao
                MontaDadosLVTeste "S", vListviewImpExpedicao
                PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "E", vListviewImpExpedicao
                PersonaColLVTeste 6, "N", "P", "", "S", "N", "N", "E", vListviewImpExpedicao
                If vListviewImpExpedicao.ListItems.Count > 0 Then ajusta_LVTeste vListviewImpExpedicao
            End If
            If apontaLV = 19 Then
                'vQdtFrom = 1
                MontaCabLV "Nº Relatório", "FCE", "Projeto", "Descrição", "Data emissão", "Status Impressão", "Tipo", "Status FCE", "", "", "", "", "", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewImpInspecao
                MontaDadosLVTeste "S", vListviewImpInspecao
                PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "E", vListviewImpInspecao
                PersonaColLVTeste 7, "N", "P", "", "S", "N", "N", "E", vListviewImpInspecao
                If vListviewImpInspecao.ListItems.Count > 0 Then ajusta_LVTeste vListviewImpInspecao
            End If
            If apontaLV = 20 Then
                'vQdtFrom = 7
                MontaCabLV "Descrição", "FCE", "Peso Líquido (FAT)", "Peso Bruto (FAT)", "Valor Bruto (FAT)", "Valor Líquido (FAT)", "Data Cadastro(FCE)", "Valor Original (FIN)", "Valor Baixado (FIN)", "Valor Receber (FIN)", "Peso (COM)", "Valor Vendido (COM)", "Status", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewFaturamentoFCE
                MontaDadosLVTeste "S", vListviewFaturamentoFCE
                PersonaColLVTeste 1, "S", "S", "", "N", "N", "N", "E", vListviewFaturamentoFCE
                PersonaColLVTeste 2, "N", "N", "", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 3, "N", "N", "", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 4, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 5, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 7, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 8, "N", "N", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 9, "S", "S", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 10, "N", "N", "", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 11, "S", "S", "R$ ", "N", "N", "S", "D", vListviewFaturamentoFCE
                PersonaColLVTeste 12, "N", "P", "", "S", "N", "N", "E", vListviewFaturamentoFCE
                If vListviewFaturamentoFCE.ListItems.Count > 0 Then ajusta_LVTeste vListviewFaturamentoFCE
            End If
            If apontaLV = 21 Then
                'vQdtFrom = 1
                MontaCabLV "Código", "Nome do usuário", "ID Setor", "Nome Setor", "ID Função", "Nome Função", "ID CC", "Nome CC", "Empresa", "D. Cadastro", "D. Contrato ini.", "D. Contrato Fim", "Ativo", "", "", "", "", "", "", "", ""
                contaColLVTeste vListviewTerceiros
                MontaDadosLVTeste "S", vListviewTerceiros
                PersonaColLVTeste 12, "N", "P", "", "S", "N", "N", "E", vListviewTerceiros
                If vListviewTerceiros.ListItems.Count > 0 Then ajusta_LVTeste vListviewTerceiros
            End If
            'montaLV1 (apontaLV)
            
            Principal.StatusBar1.Panels(5).Text = "Registros: " & vListViewPrincipal.ListItems.Count
            
            '--- A VARIAVEL ABAIXO DEFINE A ACAO A SER TOMADA EM RELAÇÃO AO BOTÃO cmdClose(Index) AO CHAMAR A FUNCTION construirBotaoClose
            'vAcaoTab = "CLOSE"
            'construirBotaoClose frmPesqGeralTeste2.SSTab1
'-- IMPRIMIR
        Case 7, 17, 27, 37, 47, 57, 67, 77, 87, 97, 107, 117, 127, 137, 147, 157, 167, 177, 187, 197
            Pesquisa = "Imprimir"
            If apontaLV = 9 Or apontaLV = 12 Then
                frmPrintRels.Show 1
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
                AlteraListview indiceVarGlobal, vListViewPrincipal
                montaDadosVendas
                frmPrintRels.Show 1
            ElseIf apontaLV = 2 Or apontaLV = 3 Or apontaLV = 5 Or apontaLV = 6 Or apontaLV = 11 Or apontaLV = 17 Then
                'FCRGeral.Show 1
            End If
        
'-- ATUALIZAR
        Case 8, 18, 28, 38, 48, 58, 68, 78, 88, 98, 108, 118, 128, 138, 148, 158, 168, 178, 188, 198
            If apontaLV = 9 Then
                AlteraListview 1, vListViewPrincipal
                
                If vRetrabalho <> "-" Then
                    Pesquisa = "editar"
                Else
                    Pesquisa = "novo"
                End If
                vTime = Time
                vTime = RemoveMask(vTime)
                frmRetrabalho.Show 1
            ElseIf apontaLV = 5 Then
                AlteraListview 1, vListViewPrincipal
                If varGlobal2 <> "-" And varGlobal2 <> "?" Then
                    frmFCE.Show 1
                ElseIf varGlobal2 = "?" Then
                    varGlobal2 = ""
                    Exit Sub
                Else
                    mobjMsg.Abrir "Nenhuma FCE selecionada", Ok, critico, "ZEUS"
                End If
            End If
        
'-- BAIXA PARCIAL | AFASTAMENTO
        Case 9, 19, 29, 39, 49, 59, 69, 79, 89, 99, 109, 119, 129, 139, 149, 159, 169, 179, 189, 199
            If apontaLV = 9 Then
                AlteraListview 2, vListViewPrincipal
                frmBaixaParcialOS.Show 1
                'mobjMsg.Abrir "Rotina de Baixa parcial de OS em desenvolvimento", Ok, informacao, "Atenção"
            ElseIf apontaLV = 5 Then
                frmImpostosServicos.Show 1
            End If
            
'-- PROGRAMACAO (VERIFICAR COMO RESOLVER O PROBLEMA DE INDICE)
        Case 10, 110, 210, 310, 410, 510, 610, 710, 810, 910, 1010, 1110, 1210, 1310, 1410, 1510, 1610, 1710, 1810, 1910
            If apontaLV = 9 Then
                'AlteraListview 1
                frmProgramacao.Show
            ElseIf apontaLV = 20 Then
                AlteraListview 2, vListViewPrincipal
                varGlobal = Mid$(varGlobal, 1, 4)
                frmAlteraStatusFCE.Show 1
            End If
    End Select
    Exit Sub
Err:
    mobjMsg.Abrir "Nenhum item selecionado", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnSubClassSSTAB SSTab1.HWnd
End Sub

Private Sub ListView2_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView2(Index), ColumnHeader
'    Select Case Index
'        Case 0
'            ColumnSort ListView2(Index), ColumnHeader
'            Combo1.Text = ColumnHeader.Text
'        Case 1
'        Case 2
'    End Select
End Sub

Public Sub ColumnSort(ListViewControl As Listview, Column As ColumnHeader)
    With ListViewControl
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
    
    If apontaLV = 5 Then PersonaColLVTeste 17, "N", "P", "", "N", "N", "N", "E", ListViewControl 'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
    If apontaLV = 9 Then PersonaColLVTeste 12, "N", "P", "", "N", "N", "N", "E", ListViewControl  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
    If apontaLV = 8 Then PersonaColLVTeste 6, "N", "P", "", "N", "N", "N", "E", ListViewControl  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
    If apontaLV = 7 Then PersonaColLVTeste 8, "N", "P", "", "N", "N", "N", "E", ListViewControl  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
    If apontaLV = 6 Then PersonaColLVTeste 13, "N", "P", "", "N", "N", "N", "E", ListViewControl  'corCol igual a 'P' significa que é para colorir a linha p identificar o Tipo de FCE
End Sub

Private Sub ListView2_DblClick(Index As Integer)
    setaComponentesTab SSTab1
    Me.MousePointer = 11
    If vEdi <> "N" Then
        Pesquisa = "editar"
        If apontaLV = 16 Or apontaLV = 17 Then
            If apontaLV = 16 Then
                vSituacao = "INSPEÇÃO DE FABRICAÇÃO"
            Else
                vSituacao = "EXPEDIÇÃO"
            End If
            AlteraListview 2, vListViewPrincipal
        Else
            AlteraListview indiceVarGlobal, vListViewPrincipal
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
                        If vListViewPrincipal.SelectedItem.ListSubItems.Item(3) = "-" Then
                            mobjMsg.Abrir "Não Exitem lançamentos financeiros para a FCE nº: " & Mid$(varGlobal, 1, 4), Ok, critico, "ZEUS"
                            Exit Sub
                        End If
                        montaDadosVendas
                    End If
                    chamaForm.Show 1
                End If
            End If
        End If
        'HabBotoes
    End If
    Me.MousePointer = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
    setaComponentesTab SSTab1
End Sub

'Private Sub SSTab1_Click(PreviousTab As Integer)
'    verificaTabAberta SSTab1
'End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    If montaLV1(apontaLV) = True Then
        Timer2.Enabled = False
        'Timer3.Enabled = True
    End If
End Sub

Private Sub Timer3_Timer()
    'construirBotaoClose frmPesqGeralTeste2.SSTab1
    'Timer3.Enabled = False
End Sub

Private Sub AlteraListview(qtdCol As Integer, vListview As Listview)
On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = vListview.ListItems.Count
    For X = 1 To Y
        If vListview.ListItems.Item(X).Selected = True Then
            'If ListView1.CheckBoxes = True Then ListView1.ListItems.Item(X).Checked = True
            Exit For
        End If
    Next
    If qtdCol = 1 Then
        varGlobal = vListview.ListItems.Item(X)
    ElseIf qtdCol = 3 Then
        varGlobal = vListview.SelectedItem.ListSubItems.Item(1)
    Else
        varGlobal = vListview.ListItems.Item(X) & vListview.SelectedItem.ListSubItems.Item(1)
    End If
    If apontaLV = 9 Then
        vRetrabalho = vListview.SelectedItem.ListSubItems.Item(9)
    End If
    If apontaLV = 5 Then
        If vListview.SelectedItem.ListSubItems.Item(16).ReportIcon = "CONCLUIDA" Then
            varGlobal2 = "?"
            mobjMsg.Abrir "FCE concluida não pode ser editada", Ok, critico, "Atenção"
        End If
        If varGlobal2 <> "?" Then varGlobal2 = vListview.SelectedItem.ListSubItems.Item(13)
    ElseIf apontaLV = 6 Then
        'SE A FCE ESTIVER CONCLUIDA NÃO CONSEGUE CRIAR NOVAS LISTAS DE MATERIAIS
        If vListview.SelectedItem.ListSubItems.Item(12).ReportIcon = "CONCLUIDA" Then
            varGlobal2 = "?"
            'mobjMsg.Abrir "FCE concluida não pode ser criado nova LM", Ok, critico, "Atenção"
        End If
    ElseIf apontaLV = 8 Then
        'SE A FCE ESTIVER CONCLUIDA NÃO CONSEGUE EDITAR A LISTA DE MATERIAIS
        If vListview.SelectedItem.ListSubItems.Item(5).ReportIcon = "CONCLUIDA" Then
            varGlobal2 = "?"
            'mobjMsg.Abrir "FCE concluida. A LM não pode ser EDITADA", Ok, critico, "Atenção"
        End If
    End If
    
    If apontaLV = 18 Then vQualquerDado(20, 1) = vListview.SelectedItem.ListSubItems.Item(1)
    
    
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
        vQualquerDado(0, 30) = "RELATÓRIO DE INSPEÇÃO DE " & vListview.SelectedItem.ListSubItems.Item(6)
    End If
    
    removeLinha = X
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        varGlobal = ""
        If vListview.ListItems.Count <> 0 Then
            mobjMsg.Abrir "Nenhum registro cadastrado ou selecionado", Ok, critico, "Atenção"
        End If
        Exit Sub
    End If
End Sub

