VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
            Object.Width           =   7532
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
   Begin MAESTRO.XTREMERibbon Ribbon 
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
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":3358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":4D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":59E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":66C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":739A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8074
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":8D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":9A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":A702
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":B3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":C0B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":CD90
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":DA6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":E744
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":F41E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":100F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":10DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":11AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":12786
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":13460
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":1413A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":14E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":15AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Principal.frx":167C8
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
         Picture         =   "Principal.frx":174A2
         Top             =   0
         Width           =   20400
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "Principal.frx":1F734
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

'Faz a imagem caber no formul�rio MDI
Private Sub MDIForm_Resize()
    On Error Resume Next
    Set Image1.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")
    EstendeImagem
End Sub

Private Sub MDIForm_Load()
On Error GoTo ErrHandler
'On Error Resume Next
'Recupera o Tema atual
Tema = GetValue(App.Path & "\CONFIG.ini", "Tema", "NomeTema", "")

'Pega o Skin atual salvo na pasta principal com o nome MySkin
Skin1.LoadSkin App.Path & "\MySkin.skn"
Skin1.ApplySkin Me.HWnd

Me.Caption = "SGC - Sistema de Gest�o em Compet�ncias" & " - Vers�o: " & App.Major & "." & App.Minor & "." & App.Revision

'Pega a imagem de funco atual salva na pasta principal com o nome PlanoDeFundo
Set Principal.Picture = LoadPicture(App.Path & "\PlanoDeFundo.jpg")

'### >> Aqui come�a o Ribbon << ############################################################
'# SET Theme BEFORE ALL
Ribbon.Theme = Tema

'# Set ImageList to use for icons
Ribbon.ImageList = ImageList3

'# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
Ribbon.ButtonCenter = False

''# Add Tabs ---ID - Caption
'Ribbon.AddTab "1", "Cadastros"
'Ribbon.AddTab "2", "Recrutamento"
'Ribbon.AddTab "3", "Decis�o"
'Ribbon.AddTab "4", "Capacita��o"
'Ribbon.AddTab "5", "Relat�rios"
'Ribbon.AddTab "6", "Configura��es"
'Ribbon.AddTab "7", "Sobre"

''# Add Cats ---ID - Tab - Caption - ShowDialogButton
''>> Tab Recrutamento - 01/10
'Ribbon.AddCat "1", "2", "Sele��o de pessoal", False

''>> Tab Decis�o - 11/20
'Ribbon.AddCat "11", "3", "Decis�es gerenciais", False

''>> Tab Capacita��o - 21/30
'Ribbon.AddCat "21", "4", "Capacita��o de pessoal", False

''>> Tab Relat�rios - 31/50
'Ribbon.AddCat "31", "5", "Relat�rios", False

''>> Tab Configura��es - 51/60
'Ribbon.AddCat "51", "6", "Parametriza��es", False
'Ribbon.AddCat "52", "6", "Apar�ncia", False

''>> Tab Sobre - 61/70
'Ribbon.AddCat "61", "7", "Sobre", False

''>> Tab Cadastro - 81/100
'Ribbon.AddCat "81", "1", "Prim�rios", False
'Ribbon.AddCat "82", "1", "Secund�rios", False

''# Add Button --- ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
''------------------------------ Tab Cadastro -----------------------------------------------
''>> Cat - Prim�rios = 81
'Ribbon.AddButton "81", "81", "Departamentos", 1
'Ribbon.AddButton "82", "81", "Setores", 2
'Ribbon.AddButton "84", "81", "Cargos", 3
'Ribbon.AddButton "85", "81", "Habilidades", 4
'Ribbon.AddButton "86", "81", "Escolaridades", 5
'Ribbon.AddButton "87", "81", "Avalia��es", 6

''>> Cat - Secund�rios = 82
'Ribbon.AddButton "88", "82", "Cursos/Treinamentos", 7
'Ribbon.AddButton "89", "82", "Matriz de Capacita��o", 8
'Ribbon.AddButton "90", "82", "Candidatos", 9
'Ribbon.AddButton "91", "82", "Colaboradores", 10

''------------------------------ Tab Recrutamento ---------------------------------------------
''>> Cat - Sele��o de pessoal = 1
'Ribbon.AddButton "1", "1", "Requisi��o de pessoal", 11
'Ribbon.AddButton "2", "1", "Processo seletivo", 12

''------------------------------ Tab Decis�o ---------------------------------------------
''>> Cat - Decis�es gerenciais = 11
'Ribbon.AddButton "11", "11", "PDO", 13

''------------------------------ Tab Capacita��o ---------------------------------------------
''>> Cat - Capacita��o de pessoal = 21
'Ribbon.AddButton "21", "21", "Programa��o", 14
'Ribbon.AddButton "22", "21", "Restri��es", 15
'Ribbon.AddButton "23", "21", "INTD", 16
'Ribbon.AddButton "24", "21", "ADP", 17

''------------------------------ Tab Relat�rios ---------------------------------------------
''>> Cat - Relat�rios = 31
'Ribbon.AddButton "31", "31", "Gr�ficos de Compet�ncias", 26
'Ribbon.AddButton "32", "31", "Programa��o anual de treinamentos", 27
'Ribbon.AddButton "33", "31", "Rela��o de cargos por treinamento", 26
'Ribbon.AddButton "34", "31", "Rel-04", 26

''------------------------------ Tab Configura��es ---------------------------------------------
''>> Cat - Parametriza��es = 51
'Ribbon.AddButton "51", "51", "Sistema", 18
'Ribbon.AddButton "52", "51", "Grupos", 19
'Ribbon.AddButton "53", "51", "Usu�rios", 20

''>> Cat - Apar�ncia = 52
'Ribbon.AddButton "54", "52", "Menu", 21
'Ribbon.AddButton "55", "52", "Skin", 22
'Ribbon.AddButton "56", "52", "Fundo", 23

''------------------------------ Tab Sobre -----------------------------------------------
''>> Cat - Sobre = 61
'Ribbon.AddButton "61", "61", "Sobre SGC", 24
'Ribbon.AddButton "62", "61", "Ajuda do SGC", 25
'---------------------------------------------------------------------------------------------

'# Repaint Ribbon
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

Private Sub Ribbon_CatClick(ByVal ID As String, ByVal Caption As String)
Select Case ID
    Case Is = 2
        'PopupMenu mnuProdutos
    Case Is = 3
        'PopupMenu mnuServi�os
    Case Is = 14
        'PopupMenu mnuEstoque
    Case Is = 31
        'PopupMenu mnuVendasCanceladas
    Case Is = 56
        'PopupMenu mnuOficina
    Case Is = 81
        'PopupMenu mnuConsultaRapida
    Case Is = 82
        'PopupMenu mnuLancamentos
    Case Is = 84
        'PopupMenu mnuAgenda
    Case Is = 85
        'PopupMenu mnuFerramentas
    Case Is = 86
        'PopupMenu mnuSobre
End Select
End Sub

Private Sub Ribbon_ButtonClick(ByVal ID As String, ByVal Caption As String)
    Pesquisa = ""
    'MeuLV.cmdconsulta(9).Visible = False
    vControlaDim = 0
    Tipo = True
    checaFiltro = True
    If ID = 81 Then '(Departamento)
        apontaLV = 2
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 82 Then '(Setores)
        apontaLV = 3
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 84 Then '(Cargos)
        apontaLV = 4
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 85 Then '(Habilidades)
        apontaLV = 5
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 86 Then '(Escolaridade)
        apontaLV = 6
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 87 Then '(Avalia��es)
        apontaLV = 11
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 88 Then '(Cursos e Treinamentos)
        apontaLV = 8
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 89 Then '(Matriz de capacita��o)
        apontaLV = 9
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 90 Then '(Candidatos)
        apontaLV = 1
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 91 Then '(Colaboradore)
        apontaLV = 0
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
'----------
    If ID = 1 Then '(Requisi��es)
        apontaLV = 7
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 2 Then '(Processo seletivo)
        apontaLV = 15
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
'----------
    If ID = 11 Then '(PDO - Processo Decis�rio Organizacional)
        apontaLV = 17
        FiltroGeral = "N�o Avaliados"
        MontaLV (apontaLV)
    End If
'----------
    If ID = 21 Then '(Programa��o)
        MeuLV.ListView1.CheckBoxes = True
        FiltroGeral = "Ativos pendentes"
        apontaLV = 10
        MontaLV (apontaLV)
        'MeuLV.ListView1.Checkboxes = False
    End If
    If ID = 22 Then '(Restri��es)
        MeuLV.ListView1.CheckBoxes = True
        FiltroGeral = "Ativos"
        apontaLV = 12
        MontaLV (apontaLV)
        'MeuLV.ListView1.Checkboxes = False
    End If
    If ID = 23 Then '(INTD - Identifica��o de Necessidade de Treinamento e Desenvolvimento)
        apontaLV = 16
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 24 Then '(ADP - Avalia��o de Desenvolvimento Pessoal)
        apontaLV = 18
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    
    
    If ID = 31 Then
        If atualizaCandidatos = False Then
            mobjMsg.Abrir "N�o h� dados suficientes para gerar os gr�ficos", Ok, critico, "Aten��o"
            Exit Sub
        Else
            criaTabTemp
            FCRGrafico.Show 1
        End If
    End If
    If ID = 32 Then
        strAno = InputBox("Informe o ano", "SGC")
        If StrPtr(strAno) = 0 Then
            mobjMsg.Abrir "Relat�rio Cancelado", Ok, critico, "Aten��o"
        Else
            If strAno <> "" Then
                FCRProgTrei.Show 1
            Else
                mobjMsg.Abrir "� necess�rio informar o ano", Ok, critico, "Aten��o"
            End If
        End If
    End If
    If ID = 33 Then
        FCRTreinCargo.Show 1
    End If
    
    If ID = 51 Then '(Sistema)
        'Principal.aicAlphaImage1.Visible = True
        Set chamaForm = New frmConfSistema
        frmConfSistema.Show 1
        'Principal.aicAlphaImage1.Visible = False
    End If
    If ID = 52 Then '(Grupos)
        apontaLV = 14
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 53 Then '(Usu�rios)
        apontaLV = 13
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
'----------
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
    If ID = 61 Then '(Sobre)
        frmRegistro.Show 1
    End If

    If ID = 62 Then '(Ajuda)
        LoadEXE (App.Path & "\SGCHHelp.exe")
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
    mobjMsg.Abrir "Ajuda em constru��o, aguarde.", , informacao, "Master System"
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
On Error GoTo erro
    Dim X As Integer
    Dim nofreeze As Integer
    X = Shell(Dir, 1)
    nofreeze = DoEvents()
    Exit Sub
erro:
    If Err.Number = 6 Then Exit Sub
    mobjMsg.Abrir "Arquivo de HELP n�o foi localizado !!! Verifique sua localiza��o ...", Ok, critico, "Aten��o"
End Sub

Private Sub montaTabMenu()
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    cnBanco.BeginTrans
    sqlDeletar = "Delete from tbMenu"
    rsDeletar.Open sqlDeletar, cnBanco
   
    'ADICIONA TABS/CATS
    SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'01','TAB','Cadastros','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'81','CAT','Prim�rios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'82','CAT','Secund�rios','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8181','BUT','Departamentos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8182','BUT','Setores','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8184','BUT','Cargos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8185','BUT','Habilidades','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8186','BUT','Escolaridade','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8187','BUT','Avalia��es','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8288','BUT','Cursos/Treinamentos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8289','BUT','Matriz de Capacita��o','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8290','BUT','Candidatos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,'8291','BUT','Colaboradores','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'02','TAB','Recrutamento','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'01','CAT','Sele��o de pessoal','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'0101','BUT','Requisi��o de pessoal','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,'0102','BUT','Processo seletivo','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'03','TAB','Decis�o','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'11','CAT','Decis�es gerenciais','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,'1111','BUT','PDO','" & vCodcoligada & "')Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'04','TAB','Capacita��o','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'21','CAT','Capacita��o de pessoal','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'2121','BUT','Programa��o','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'2122','BUT','Restri��es','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'2123','BUT','INTD','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,'2124','BUT','ADP','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'05','TAB','Relat�rios','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'31','CAT','Relat�rios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'3131','BUT','GC','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'3132','BUT','PAT','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'3133','BUT','RCT','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,'3134','BUT','Rel-04','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'06','TAB','Configura��es','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'51','CAT','Parametriza��es','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'52','CAT','Apar�ncia','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5151','BUT','Sistema','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,'5152','BUT','Grupos','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,5153,'BUT','Usu�rios','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,5254,'BUT','Menu','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,5255,'BUT','Skin','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(6,5256,'BUT','Fundo','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,07,'TAB','Sobre','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,61,'CAT','Sobre','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,6161,'BUT','Sobre SGC','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(7,6162,'BUT','Ajuda do SGC','" & vCodcoligada & "');"
    rsMenu.Open SqlMenu, cnBanco
    cnBanco.CommitTrans
    Set rsMenu = Nothing
End Sub

Private Sub abreConfMenu()
'    SqlConf = "Select * from tbconfgrupo Where tbconfgrupo.codcoligada = '" & vCodcoligada & "' and tbconfgrupo.idgrupo = '" & XCodGrp & "'order by id"
    SqlConf = "Select * from tbconfgrupo Where tbconfgrupo.idgrupo = '" & XCodGrp & "'order by id"
    rsConf.Open SqlConf, cnBanco, adOpenKeyset, adLockReadOnly
End Sub

Private Sub fechaConfMenu()
    rsConf.Close
    Set rsConf = Nothing
End Sub

Private Sub montaMenu()
    While Not rsConf.EOF
    If rsConf.Fields(5) = "S" Then
        If rsConf.Fields(3) = "TAB" Then
            Ribbon.AddTab rsConf.Fields(1), rsConf.Fields(4)
        End If
        If rsConf.Fields(3) = "CAT" Then
            Ribbon.AddCat rsConf.Fields(2), rsConf.Fields(1), rsConf.Fields(4), False
        End If
        If rsConf.Fields(3) = "BUT" Then
            Ribbon.AddButton Mid$(rsConf.Fields(2), 3, 2), Mid$(rsConf.Fields(2), 1, 2), rsConf.Fields(4), rsConf.Fields(8)
        End If
        If rsConf.Fields(3) = "CHK" Then
            If rsConf.Fields(4) = "CHKINC" Then vInc = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKEDI" Then vEdi = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKSAL" Then vSal = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKEXC" Then vExc = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKIMP" Then vImp = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKFIL" Then vFil = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKAVA" Then vAva = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKADI" Then vAdi = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKDEM" Then vDem = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKADIRES" Then vAdiRes = rsConf.Fields(5)
            If rsConf.Fields(4) = "CHKADIREP" Then vAdiRep = rsConf.Fields(5)
        End If
    End If
    rsConf.MoveNext
    Wend
    Ribbon.Refresh
End Sub

Private Function atualizaCandidatos()
On Error Resume Next
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
End Function

Private Sub GravaColaboradores()
On Error Resume Next
    Dim rsGravaColaboradores As New ADODB.Recordset
    Dim sqlGravaColaboradores As String
    Dim vIdent As Integer
    vIdent = rsCandidatos.Fields(0)
    
    sqlGravaColaboradores = "INSERT INTO ##Tempglobal(id,cpf,nomecolaborador,departamento,setor,experiencia,habilidade,treinamento,formacao) VALUES('" & rsCandidatos.Fields(0) & "','" & rsCandidatos.Fields(1) & "','" & rsCandidatos.Fields(2) & "','" & rsCandidatos.Fields(3) & "','" & rsCandidatos.Fields(4) & "','" & Replace(RemoveMask(Label37), ",", ".") & "','" & Replace(RemoveMask(Label38), ",", ".") & "','" & Replace(RemoveMask(Label39), ",", ".") & "','" & Replace(RemoveMask(Label41), ",", ".") & "')"
    rsGravaColaboradores.Open sqlGravaColaboradores, cnBanco
End Sub

Private Sub criaTabTemp()
On Error Resume Next
    'Criando uma tabela tempor�ria global
    Dim rsTabTemp As New ADODB.Recordset
    Dim SqlTabTemp As String
    SqlTabTemp = "CREATE TABLE ##Tempglobal(id INT NOT NULL,CPF VARCHAR(50) NOT NULL,nomecolaborador VARCHAR(100) NOT NULL,departamento VARCHAR(100) NOT NULL, setor VARCHAR(100) NOT NULL, experiencia FLOAT NOT NULL, habilidade FLOAT NOT NULL, treinamento FLOAT NOT NULL, formacao FLOAT NOT NULL)"
    rsTabTemp.Open SqlTabTemp, cnBanco
End Sub

