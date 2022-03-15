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
   Begin MSComctlLib.ImageList ImgList3 
      Left            =   5640
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":0000
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":C932
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":19264
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":25B96
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":324C8
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3EDFA
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4B72C
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5805E
            Key             =   "AGUARDE-03"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":64990
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":712C2
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7DBF4
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8A526
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":96E58
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A378A
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B00BC
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":BC9EE
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":C9320
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":D5C52
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":E2584
            Key             =   "DUVIDA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList2 
      Left            =   5040
      Top             =   8400
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
            Picture         =   "frmPesqGeralTeste2.frx":EEEB6
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":FB7E8
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":10811A
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":114A4C
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":12137E
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":12DCB0
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":13A5E2
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":146F14
            Key             =   "AGUARDE-03"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":153846
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":153CE0
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":15417A
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":154614
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":154AAE
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":154F48
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1553E2
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":15587C
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1621AE
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":16EAE0
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":17B412
            Key             =   "DUVIDA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList1 
      Left            =   4440
      Top             =   8400
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
            Picture         =   "frmPesqGeralTeste2.frx":187D44
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":194676
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1A0FA8
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1B28BB
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1BF1ED
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1CDD78
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1DA6AA
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1E6FDC
            Key             =   "AGUARDE-03"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F390E
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F3DA8
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F4242
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F46DC
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F4B76
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F5010
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F54AA
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1F5944
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":208578
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":214EAA
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2217DC
            Key             =   "DUVIDA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5640
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   56
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":22E10E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":23AA40
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":247372
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":253CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2605D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":26CF08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":27983A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":28616C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":292A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":29F3D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2ABD02
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2B8634
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2C4F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2D1898
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2DE1CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2EAAFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2F742E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":303D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":310692
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":31CFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3298F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":336228
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":342B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":34F48C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":35BDBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3686F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":375022
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":381954
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":38E286
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":39ABB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3A74EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3B3E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3C074E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3CD080
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3D99B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3E62E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3F2C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3FF548
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":40BE7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4187AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4250DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":431A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":431CD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":43E605
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":44AF37
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":457869
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":46419B
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":470ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":47D3FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":489D31
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":496663
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4A2F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4AF8C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4BC1F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4C8B2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4D545D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5040
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   56
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4E1D8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4EE6C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4FAFF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":507925
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":514257
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":520B89
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":52D4BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":539DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":54671F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":553051
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":55F983
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":56C2B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":578BE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":585519
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":591E4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":59E77D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5AB0AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5B79E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5C4313
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5D0C45
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5DD577
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5E9EA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":5F67DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":60310D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":60FA3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":61C371
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":628CA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6355D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":641F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":64E839
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":65B16B
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":667A9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6743CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":680D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":68D633
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":699F65
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6A6897
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6B31C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6BFAFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6CC42D
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6D8D5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6E5691
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6E5954
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6F2286
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":6FEBB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":70B4EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":717E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":72474E
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":731080
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":73D9B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":74A2E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":756C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":763548
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":773552
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":77FE84
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":78C7B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   72
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7990E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7A8AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7B5412
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7C1D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7D108A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7DD9BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7ECED1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":7FAEDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":80780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":814140
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":820A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":82D3A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":839CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":846608
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":852F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":85F86C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":86F833
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":87C165
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":888A97
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8953C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8A33D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8AFD06
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8C0154
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8CCA86
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8D93B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8E5CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8F5CB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9025E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":90EF15
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":91B847
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":928179
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":934AAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":94455D
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9530E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":95FA1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":96C34C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":978C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9855B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":991EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":99E814
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9AB146
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9B7A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9B7D3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9C466D
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9D8DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9E56EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9F2020
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9FE952
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A0B284
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A17BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A2A7EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A3711C
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A4AFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A5790C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A6423E
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A70B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A7D4A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A89DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A96706
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AA3038
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AAF96A
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":ABC29C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AC8BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AD5500
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AE1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AEE764
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":AFB096
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B079C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B142FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B20C2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B3107A
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B3D9AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8040
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
      Left            =   7560
      Top             =   8400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6960
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
      TabPicture(0)   =   "frmPesqGeralTeste2.frx":B4F8A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView2(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClose(40)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Carregando..."
      TabPicture(1)   =   "frmPesqGeralTeste2.frx":B4F8BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Carregando..."
      TabPicture(2)   =   "frmPesqGeralTeste2.frx":B4F8DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Carregando..."
      TabPicture(3)   =   "frmPesqGeralTeste2.frx":B4F8F6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Carregando..."
      TabPicture(4)   =   "frmPesqGeralTeste2.frx":B4F912
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Carregando..."
      TabPicture(5)   =   "frmPesqGeralTeste2.frx":B4F92E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Carregando..."
      TabPicture(6)   =   "frmPesqGeralTeste2.frx":B4F94A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Carregando..."
      TabPicture(7)   =   "frmPesqGeralTeste2.frx":B4F966
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Carregando..."
      TabPicture(8)   =   "frmPesqGeralTeste2.frx":B4F982
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Carregando..."
      TabPicture(9)   =   "frmPesqGeralTeste2.frx":B4F99E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Carregando..."
      TabPicture(10)  =   "frmPesqGeralTeste2.frx":B4F9BA
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Carregando..."
      TabPicture(11)  =   "frmPesqGeralTeste2.frx":B4F9D6
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Carregando..."
      TabPicture(12)  =   "frmPesqGeralTeste2.frx":B4F9F2
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "Carregando..."
      TabPicture(13)  =   "frmPesqGeralTeste2.frx":B4FA0E
      Tab(13).ControlEnabled=   0   'False
      Tab(13).ControlCount=   0
      TabCaption(14)  =   "Carregando..."
      TabPicture(14)  =   "frmPesqGeralTeste2.frx":B4FA2A
      Tab(14).ControlEnabled=   0   'False
      Tab(14).ControlCount=   0
      TabCaption(15)  =   "Carregando..."
      TabPicture(15)  =   "frmPesqGeralTeste2.frx":B4FA46
      Tab(15).ControlEnabled=   0   'False
      Tab(15).ControlCount=   0
      TabCaption(16)  =   "Carregando..."
      TabPicture(16)  =   "frmPesqGeralTeste2.frx":B4FA62
      Tab(16).ControlEnabled=   0   'False
      Tab(16).ControlCount=   0
      TabCaption(17)  =   "Carregando..."
      TabPicture(17)  =   "frmPesqGeralTeste2.frx":B4FA7E
      Tab(17).ControlEnabled=   0   'False
      Tab(17).ControlCount=   0
      TabCaption(18)  =   "Carregando..."
      TabPicture(18)  =   "frmPesqGeralTeste2.frx":B4FA9A
      Tab(18).ControlEnabled=   0   'False
      Tab(18).ControlCount=   0
      TabCaption(19)  =   "Carregando..."
      TabPicture(19)  =   "frmPesqGeralTeste2.frx":B4FAB6
      Tab(19).ControlEnabled=   0   'False
      Tab(19).ControlCount=   0
      TabCaption(20)  =   "Carregando..."
      TabPicture(20)  =   "frmPesqGeralTeste2.frx":B4FAD2
      Tab(20).ControlEnabled=   0   'False
      Tab(20).ControlCount=   0
      TabCaption(21)  =   "Carregando..."
      TabPicture(21)  =   "frmPesqGeralTeste2.frx":B4FAEE
      Tab(21).ControlEnabled=   0   'False
      Tab(21).ControlCount=   0
      TabCaption(22)  =   "Carregando..."
      TabPicture(22)  =   "frmPesqGeralTeste2.frx":B4FB0A
      Tab(22).ControlEnabled=   0   'False
      Tab(22).ControlCount=   0
      TabCaption(23)  =   "Carregando..."
      TabPicture(23)  =   "frmPesqGeralTeste2.frx":B4FB26
      Tab(23).ControlEnabled=   0   'False
      Tab(23).ControlCount=   0
      TabCaption(24)  =   "Carregando..."
      TabPicture(24)  =   "frmPesqGeralTeste2.frx":B4FB42
      Tab(24).ControlEnabled=   0   'False
      Tab(24).ControlCount=   0
      TabCaption(25)  =   "Carregando..."
      TabPicture(25)  =   "frmPesqGeralTeste2.frx":B4FB5E
      Tab(25).ControlEnabled=   0   'False
      Tab(25).ControlCount=   0
      TabCaption(26)  =   "Carregando..."
      TabPicture(26)  =   "frmPesqGeralTeste2.frx":B4FB7A
      Tab(26).ControlEnabled=   0   'False
      Tab(26).ControlCount=   0
      TabCaption(27)  =   "Carregando..."
      TabPicture(27)  =   "frmPesqGeralTeste2.frx":B4FB96
      Tab(27).ControlEnabled=   0   'False
      Tab(27).ControlCount=   0
      TabCaption(28)  =   "Carregando..."
      TabPicture(28)  =   "frmPesqGeralTeste2.frx":B4FBB2
      Tab(28).ControlEnabled=   0   'False
      Tab(28).ControlCount=   0
      TabCaption(29)  =   "Carregando..."
      TabPicture(29)  =   "frmPesqGeralTeste2.frx":B4FBCE
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
         MICON           =   "frmPesqGeralTeste2.frx":B4FBEA
         PICN            =   "frmPesqGeralTeste2.frx":B4FC06
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
      Left            =   3840
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   56
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B50022
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B50CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B519D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B526B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5338A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B54064
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B54D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B55A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B566F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B573CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B580A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B58D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B59A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5A734
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5B40E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5C0E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5CDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5DA9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5E776
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B5F450
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6012A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B60E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B61ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B627B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B63492
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6416C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B64E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B65B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B667FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B674D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B67C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B68928
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B69602
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6A2DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6AFB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6BC90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6C96A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6D644
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6E31E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6EFF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6F83A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B6FEB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B70176
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B70E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B7151D
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B721F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B72ED1
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B73BAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B74885
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B7555F
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B76239
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B76F13
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B77BED
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B788C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B78E61
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B79B3B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3840
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B7A815
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B7B227
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B7BC39
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B80653
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B8506D
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B933F6
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B940D0
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B94DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B95A84
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9675E
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B97438
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B98112
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B98DEC
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B99AC6
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9A7A0
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9B47A
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9C154
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9CB66
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9D840
            Key             =   "DUVIDA"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9E51A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B9F1F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8760
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
            ElseIf apontaLV = 5 Then
                frmReceitasDespesas.Show 1
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

