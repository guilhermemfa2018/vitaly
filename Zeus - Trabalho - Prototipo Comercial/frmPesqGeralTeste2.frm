VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPesqGeralTeste2 
   BorderStyle     =   0  'None
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17490
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
   ScaleHeight     =   10020
   ScaleWidth      =   17490
   ShowInTaskbar   =   0   'False
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
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4042
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":4D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":59F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":73AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8084
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":8D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":9A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":A712
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":B3EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":C0C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":CDA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":DA7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":E754
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":F42E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":10108
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":10DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":11ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":12796
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":13470
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1414A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":14E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":15AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":167D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":174B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":17C2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":18906
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":195E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1A2BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1AF94
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1BC6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1C948
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1D622
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1E2FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":1EFD6
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
            Picture         =   "frmPesqGeralTeste2.frx":1FCB0
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":206C2
            Key             =   "EXC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":210D4
            Key             =   "POSITIVO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":25AEE
            Key             =   "NEGATIVO"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":2A508
            Key             =   "ARQUIVADO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":38891
            Key             =   "AGUARDE-01"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3956B
            Key             =   "AGUARDE-02"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3A245
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3AF1F
            Key             =   "PENDENTE12"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3BBF9
            Key             =   "AVALIANDO1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3C8D3
            Key             =   "CONCLUIDO1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3D5AD
            Key             =   "PRETO"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3E287
            Key             =   "PENDENTE1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3EF61
            Key             =   "FABRICANDO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":3FC3B
            Key             =   "FECHADO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":40915
            Key             =   "ANDAMENTO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":415EF
            Key             =   "CONCLUIDA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":42001
            Key             =   "PARALIZADA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesqGeralTeste2.frx":42CDB
            Key             =   "DUVIDA"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   975
      Left            =   1800
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   11
      Tab             =   10
      TabsPerRow      =   11
      TabHeight       =   520
      TabPicture(0)   =   "frmPesqGeralTeste2.frx":439B5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabPicture(1)   =   "frmPesqGeralTeste2.frx":439D1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabPicture(2)   =   "frmPesqGeralTeste2.frx":439ED
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabPicture(3)   =   "frmPesqGeralTeste2.frx":43A09
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabPicture(4)   =   "frmPesqGeralTeste2.frx":43A25
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabPicture(5)   =   "frmPesqGeralTeste2.frx":43A41
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabPicture(6)   =   "frmPesqGeralTeste2.frx":43A5D
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabPicture(7)   =   "frmPesqGeralTeste2.frx":43A79
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabPicture(8)   =   "frmPesqGeralTeste2.frx":43A95
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabPicture(9)   =   "frmPesqGeralTeste2.frx":43AB1
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabPicture(10)  =   "frmPesqGeralTeste2.frx":43ACD
      Tab(10).ControlEnabled=   -1  'True
      Tab(10).ControlCount=   0
   End
End
Attribute VB_Name = "frmPesqGeralTeste2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private WithEvents ctlDynamic As VBControlExtender
Option Explicit
'Dim mo_Events As Collection
'
'Private objText(19, 15) As TextBox
'Private objFrame(19, 15) As Frame
'Private objCombo(19, 15) As ComboBox
'Private objLabel(19, 15) As Label
'Private objListview(19, 15) As MSComctlLib.Listview
''Private objButton1(19, 15) As VBControlExtender
''Private objButton(19, 15) As VBControlExtender
'Private objPicture(19, 15) As PictureBox
'Private objImage As Image
'Private vFramePrincipal As Frame
'Private vListViewPrincipal As Listview
Private vSSTab As SSTab


Private Sub Command1_Click()
    constroiTabs
    DimensionaLV1 "Métodos e Processos", vFramePrincipal, vListViewPrincipal
End Sub

Private Sub Command2_Click()
    'descontruirControles SSTab1.Tab
    statusDados SSTab1.Tab, False
    SSTab1.TabVisible(SSTab1.Tab) = False
End Sub

Private Function statusDados(vTabAtiva As Integer, VouF As Boolean)
    If vTabAtiva = 0 Then
        'Frame1.Visible = VouF
    End If
    If vTabAtiva = 1 Then
        'Frame4.Visible = VouF
    End If
    If vTabAtiva = 2 Then
        'Frame7.Visible = VouF
    End If
    If vTabAtiva = 3 Then
        'Frame10.Visible = VouF
    End If
    If vTabAtiva = 4 Then
        'Frame13.Visible = VouF
    End If
    If vTabAtiva = 5 Then
        'Frame16.Visible = VouF
    End If
    If vTabAtiva = 6 Then
        'Frame19.Visible = VouF
    End If
    If vTabAtiva = 7 Then
        'Frame22.Visible = VouF
    End If
    If vTabAtiva = 8 Then
        'Frame25.Visible = VouF
    End If
    If vTabAtiva = 9 Then
        'Frame28.Visible = VouF
    End If
    If vTabAtiva = 10 Then
        'Frame31.Visible = VouF
    End If
End Function

'Private Function constroiTabs()
'    Dim vProximaTab As Integer, X As Integer
'    X = 10
'    For vProximaTab = 0 To X
'        If SSTab1.TabVisible(vProximaTab) = False Then
'            Exit For
'        Else
'        End If
'    Next
'    If vProximaTab <= 10 Then
'        SSTab1.TabVisible(vProximaTab) = True
'        SSTab1.Tab = vProximaTab
'        construirControles vProximaTab
'        construirBotoes vProximaTab, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg", 360, 120, 615, 615, "Novo"
'        construirBotoes vProximaTab, 2, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\editar_UP.jpg", 360, 720, 615, 615, "Editar"
'        construirBotoes vProximaTab, 3, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\excluir_UP.jpg", 360, 1320, 615, 615, "Excluir"
'        construirBotoes vProximaTab, 4, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\sair_UP.jpg", 360, 1920, 615, 615, "Sair"
'        construirBotoes vProximaTab, 5, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\admitir_UP.jpg", 360, 8040, 615, 615, "Admitir"
'        construirBotoes vProximaTab, 6, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\filtro_UP.jpg", 360, 8640, 615, 615, "Filtrar"
'        construirBotoes vProximaTab, 7, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\imprimir_UP.jpg", 360, 9240, 615, 615, "Imprimir"
'        construirBotoes vProximaTab, 8, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\atualiza_UP.jpg", 360, 9840, 615, 615, "Atualizar"
'        construirBotoes vProximaTab, 9, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\afastado_UP.jpg", 360, 10440, 615, 615, "Afastamento"
'        construirBotoes vProximaTab, 10, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\prog_UP.jpg", 360, 11040, 615, 615, "Programação"
'    End If
'    statusDados vProximaTab, True
'End Function

'Private Function construirControles(vTab As Integer)
'    Set objFrame(vTab, 0) = Controls.Add("VB.Frame", "Frame1" + Trim(Str(vTab)), SSTab1)
'    With objFrame(vTab, 0)
'        .Visible = True
'        .Top = 360
'        .Left = 120
'        .Width = 16695
'        .Height = 9015
'        .Caption = "Informações"
'    End With
'    Set vFramePrincipal = objFrame(vTab, 0)
'
'    Set objFrame(vTab, 1) = Controls.Add("VB.Frame", "Frame0" + Trim(Str(vTab)), objFrame(vTab, 0))
'    With objFrame(vTab, 1)
'        .Visible = True
'        .Top = 240
'        .Left = 2760
'        .Width = 5175
'        .Height = 735
'        .Caption = "Pesquisa"
'    End With
'
'    Set objPicture(vTab, 0) = Controls.Add("VB.PictureBox", "picBg" + Trim(Str(vTab)), objFrame(vTab, 0))
'    With objPicture(vTab, 0)
'        .Visible = False
'        .Top = 360
'        .Left = 15600
'        .Width = 855
'        .Height = 495
'    End With
'
'
'    Set objFrame(vTab, 2) = Controls.Add("VB.Frame", "Frame3" + Trim(Str(vTab)), objFrame(vTab, 0))
'    With objFrame(vTab, 2)
'        .Visible = True
'        .Top = 120
'        .Left = 12360
'        .Width = 3975
'        .Height = 855
'        .Caption = "Filtro "
'        .Appearance = 0
'        .BackColor = &H8000000F
'    End With
'
'    Set objLabel(vTab, 0) = Controls.Add("VB.Label", "Label1" + Trim(Str(vTab)), objFrame(vTab, 2))
'    With objLabel(vTab, 0)
'        .Visible = True
'        .Top = 240
'        .Left = 120
'        .Width = 735
'        .Height = 255
'        .Caption = "Status: "
'    End With
'
'    Set objLabel(vTab, 1) = Controls.Add("VB.Label", "Label3" + Trim(Str(vTab)), objFrame(vTab, 2))
'    With objLabel(vTab, 1)
'        .Visible = True
'        .Top = 480
'        .Left = 120
'        .Width = 855
'        .Height = 255
'        .Caption = "Período: "
'    End With
'
'    Set objLabel(vTab, 2) = Controls.Add("VB.Label", "Label2" + Trim(Str(vTab)), objFrame(vTab, 2))
'    With objLabel(vTab, 2)
'        .Visible = True
'        .Top = 240
'        .Left = 960
'        .Width = 2055
'        .Height = 255
'        .Caption = "-"
'    End With
'
'    Set objLabel(vTab, 3) = Controls.Add("VB.Label", "Label4" + Trim(Str(vTab)), objFrame(vTab, 2))
'    With objLabel(vTab, 3)
'        .Visible = True
'        .Top = 480
'        .Left = 960
'        .Width = 2055
'        .Height = 255
'        .Caption = "-"
'    End With
'
'    Set objListview(vTab, 0) = Controls.Add("MSComctlLib.ListViewCtrl.2", "Listview2" + Trim(Str(vTab)), objFrame(vTab, 0))
'    With objListview(vTab, 0)
'        .Visible = True
'        .Top = 1080
'        .Left = 120
'        .Width = 16455
'        .Height = 7695
'        .Gridlines = True
'        .FullRowSelect = True
'        .LabelEdit = lvwManual
'        .LabelWrap = True
'        .SortKey = 0
'        .SortOrder = lvwAscending
'        .View = lvwReport
'        .BackColor = &H80000018
'        .ForeColor = &H800000
'    End With
'    Set vListViewPrincipal = objListview(vTab, 0)
'End Function

'Private Function construirBotoes(vTab As Integer, vBotao As Integer, vCaminho As String, vTop As Integer, vLeft As Integer, vWidth As Integer, vHeight As Integer, vTag As String)
''On Error Resume Next
'    Set objImage = Me.Controls.Add("VB.Image", "objImage" & vTab & vBotao, objFrame(vTab, 0))
'    With objImage
'        .Visible = True
'        .Top = vTop
'        .Left = vLeft
'        .Width = vWidth
'        .Height = vHeight
'        .Picture = LoadPicture(vCaminho)
'        .Tag = vTag
'        .ToolTipText = vTag & vTab & vBotao
'    End With
'    mo_Events.Add New cEvents
'    mo_Events(Val(vTab & vBotao)).Add_Image objImage, Val(vTab & vBotao)
'End Function

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
        'Me.Controls.Remove objButton1(vTab, i).Name
        Me.Controls.Remove objPicture(vTab, i).Name
    Next
End Function

'''Private Function desconstroiTabs()
'''    Dim i As Long
'''    For i = 0 To 10
'''        SSTab1.TabVisible(i) = False
'''    Next
'''End Function

Private Sub Form_Load()
'    vPosAtual = 1
    Set mo_Events = New Collection
    Set vSSTab = SSTab1
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
'    MudaPropPicture 'Configura Picture para colorir as linhas do listview de acordo com o Tipo de FCE
    'configControles
    desconstroiTabs vSSTab
    'constroiTabs vSSTab
    'DimensionaLV1 "Métodos e Processos", vFramePrincipal, vListViewPrincipal
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Public Sub ImgClick(p_idx As Long)
    Select Case p_idx
        Case 1
            construirBotoes 1, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg", 360, 120, 615, 615, "Novo"
        Case 4, 14, 24, 34, 44
            If contaTabsAbertas(SSTab1) = 1 Then
                Unload Me
            Else
                SSTab1.TabVisible(SSTab1.Tab) = False
            End If
    End Select
End Sub

Public Sub ImgMouseDown(p_idx As Long)
    'Msgbox "Botão: " & p_idx
    Select Case p_idx
        Case 1
            'construirBotoes 1, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_DOWN.jpg", 360, 120, 615, 615, "Novo"
    End Select
End Sub

'Public Sub ImgMouseUp(p_idx As Long)
'    'Msgbox "Botão: " & p_idx
'    Select Case p_idx
'        Case 1
'            'construirBotoes 1, 1, "D:\DESENVOLVIMENTO_VITALY\DESENVOLVIMENTO\PROJETOS\Zeus - Trabalho\Icones\Botoes\novo_UP.jpg", 360, 120, 615, 615, "Novo"
'    End Select
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Set mo_Events = Nothing
'End Sub
