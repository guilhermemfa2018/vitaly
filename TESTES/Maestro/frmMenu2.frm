VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Begin VB.Form frmMenu2 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "SGCH - Sistema de Gestão de Competência e Habilidade"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   10560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMenu2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10560
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3360
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":0E5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":1B37
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":2811
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":2D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":32A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":3F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":4C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":5930
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":660A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":72E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":7FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":8C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":9972
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":A64C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":B326
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":C000
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":CCDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":D9B4
            Key             =   ""
            Object.Tag             =   "Identificação das Necessidades de Treinamento e Desenvolvimento"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":E68E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":F368
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":10042
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":10D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu2.frx":10EAF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SGCH.ACPRibbon ACPRibbon1 
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   18627
      _ExtentY        =   3757
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8130
      Width           =   10560
      _ExtentX        =   18627
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
            Object.Width           =   4304
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
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   7995
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   1650
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   16620
      _ExtentX        =   29316
      _ExtentY        =   2910
      Image           =   "frmMenu2.frx":11B89
      Opacity         =   60
      GrayScale       =   0
      Props           =   5
      ShadowOpacity   =   90
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1440
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "frmMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Theme As Integer
Private rsConf As New ADODB.Recordset
Private SqlConf As String
Private vFechar As Integer

Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)
    Pesquisa = ""
    MeuLV.cmdconsulta(9).Visible = False
    vControlaDim = 0
    TiPo = True
    checaFiltro = True
    If ID = 0 Then '(Colaboradores)
        apontaLV = 0
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 1 Then '(Candidatos)
        apontaLV = 1
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 2 Then '(Departamentos)
        apontaLV = 2
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 3 Then '(Setores)
        apontaLV = 3
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 4 Then '(Cargos)
        apontaLV = 4
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 5 Then '(Habilidades)
        apontaLV = 5
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 6 Then '(Escolares)
        apontaLV = 6
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 7 Then '(Avaliação do treinamento)
        apontaLV = 11
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    
    If ID = 8 Then '(Requisições)
        apontaLV = 7
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 9 Then '(Processo Seletivo)
        apontaLV = 15
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 10 Then '(Treinamentos)
        apontaLV = 8
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 11 Then '(Matrizes)
        apontaLV = 9
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    
    If ID = 12 Then '(INTD)
        apontaLV = 16
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    
    If ID = 13 Then '(Programação)
        MeuLV.ListView1.CheckBoxes = True
        FiltroGeral = "Ativos pendentes"
        apontaLV = 10
        MontaLV (apontaLV)
        MeuLV.ListView1.CheckBoxes = False
    End If
        
    If ID = 14 Then '(Reprovados)
        MeuLV.ListView1.CheckBoxes = True
        FiltroGeral = "Ativos"
        apontaLV = 12
        MontaLV (apontaLV)
        MeuLV.ListView1.CheckBoxes = False
    End If
        
    If ID = 15 Then '(ADP)
        apontaLV = 18
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
        
    If ID = 16 Then '(Usuários)
        apontaLV = 13
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
    End If
    If ID = 17 Then '(Grupos)
        apontaLV = 14
        FiltroGeral = "Ativos"
        MontaLV (apontaLV)
        'frmMenu2.aicAlphaImage1.Visible = True
        'frmGrupos.Show 1
        'frmMenu2.aicAlphaImage1.Visible = False
    End If
    If ID = 18 Then ' (Sistema)
        frmMenu2.aicAlphaImage1.Visible = True
        frmConfSistema.Show 1
        frmMenu2.aicAlphaImage1.Visible = False
    End If
    If ID = 19 Then '(PDO)
        apontaLV = 17
        FiltroGeral = "Não Avaliados"
        MontaLV (apontaLV)
        'MsgBox "Em desenvolvimento"
        'frmPDO.Show 1
    End If
    If ID = 20 Then
        'Muda o tema da tela Principal
        'MudaTema
        frmRegistro.Show 1
    End If
    If ID = 21 Then
        'vHelp = WinHelp(HWnd, App.HelpFile, HELP_INDEX, CLng(0))
        LoadEXE (App.Path & "\SGCHHelp.exe")
    End If
End Sub

Private Sub MudaTema()
    Theme = Theme + 1
    'If Theme = 3 Then Theme = 0
    '# Set Theme
    ACPRibbon1.Theme = Theme
    '# Refresh control
    ACPRibbon1.Refresh
    
    '# OPTIONAL - Load Background for Form.
    'Image1.Picture = ACPRibbon1.LoadBackground
    
    '# OPTIONAL - Load Background for Form
    'frmMenu2.BackColor = ACPRibbon1.BackColor
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
    MsgBox "Arquivo de HELP não foi localizado !!! Verifique sua localização ...", vbExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        LoadEXE (App.Path & "\SGCHHelp.exe")
    End If
End Sub

Private Sub Form_Load()
    'App.HelpFile = App.Path & "\help\" & "SGCH Help.hlp"
    
    dataFilter1 = "01/01/" & DatePart("yyyy", Date)
    dataFilter2 = "31/12/" & DatePart("yyyy", Date)
    
    frmMenu2.Caption = frmMenu2.Caption & " - Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    Theme = 0

    '# SET Theme
    ACPRibbon1.Theme = Theme    ' 0 - Black
                            ' 1 - Blue
                            ' 2 - Silver

    '# OPTIONAL - Load Background for Form.
    Image1.Left = 0
    Image1.Top = ACPRibbon1.Height
    'Image1.Picture = ACPRibbon1.LoadBackground

    '# OPTIONAL - Load Background for Form
    'frmMenu2.BackColor = ACPRibbon1.BackColor

    '# Set ImageList to use for icons
    ACPRibbon1.ImageList = ImageList1


    '# Define  a exibição do botão circular e o icone do menu ACPRibbon1
    ACPRibbon1.Icon = 20

    '# Exibe o Caption do ACPRibbon1
    ACPRibbon1.Caption = Me.Caption


'# Show Button to Customize Menu
ACPRibbon1.ShowCustomMenu = True

'# Add TopButtons ---   ID - Capt. - Icons
ACPRibbon1.AddTopButton "1", "Imprimir", 22
'ACPRibbon1.AddTopButton "2", "Open", 3
'ACPRibbon1.AddTopButton "3", "Print", 4
'ACPRibbon1.AddTopButton "4", "Save", 5



'# Add TopButtons ---   ID - Capt. - Icons


    '# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
    'ACPRibbon1.ButtonCenter = False

    'ABRE TABELA TBCONFGRUPO PARA CONFERENCIA DE AUTORIZAÇÃO
    abreConfMenu
    montaMenu
    fechaConfMenu
'    'Adiciona as TABS no MENU
'    '# Add Tabs ---   ID - Caption
'    ACPRibbon1.AddTab "1", "Cadastros"
'    ACPRibbon1.AddTab "2", "Recrutamento"
'    ACPRibbon1.AddTab "3", "Capacitação"
'    ACPRibbon1.AddTab "4", "Configurações"
'    ACPRibbon1.AddTab "5", "Sobre"

'    'Cria os GRUPOS das TABS
'    '# Add Cats ---   ID - Tab - Caption - ShowDialogButton
'    ACPRibbon1.AddCat "1", "1", "Cadastros gerais", False
'    ACPRibbon1.AddCat "2", "1", "Tabelas Auxiliares", True
'    ACPRibbon1.AddCat "3", "2", "Seleção de pessoal", True
'    ACPRibbon1.AddCat "4", "3", "Capacitação de pessoal", False
'    ACPRibbon1.AddCat "5", "4", "Parametrizações", False
'    ACPRibbon1.AddCat "6", "5", "Ajuda", True

'    'Cria os BOTOES das TABS
'    '# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
'    ACPRibbon1.AddButton "0", "1", "Colaboradores" & vbNewLine, 2
'    ACPRibbon1.AddButton "1", "1", "Candidatos", 1
'    ACPRibbon1.AddButton "2", "1", "Departamentos", 3
'    ACPRibbon1.AddButton "3", "1", "Setores" & vbNewLine, 4
'    ACPRibbon1.AddButton "4", "1", "Cargos" & vbNewLine, 5
'    ACPRibbon1.AddButton "5", "2", "Habilidades funcionais", 8
'    ACPRibbon1.AddButton "6", "2", "Formação escolar", 9
'    ACPRibbon1.AddButton "7", "2", "Avaliação do treinamento", 16
'    ACPRibbon1.AddButton "8", "3", "Requisição de pessoal", 14
'    ACPRibbon1.AddButton "9", "3", "Processo seletivo", 15
'    ACPRibbon1.AddButton "10", "4", "Cursos/treinamentos", 6
'    ACPRibbon1.AddButton "11", "4", "Matriz de capacitação", 7
'    ACPRibbon1.AddButton "12", "4", "Programação", 13
'    ACPRibbon1.AddButton "13", "4", "Restrições", 17
'    ACPRibbon1.AddButton "14", "5", "Usuários", 10
'    ACPRibbon1.AddButton "15", "5", "Grupos", 11
'    ACPRibbon1.AddButton "16", "5", "Sistema", 12
'    ACPRibbon1.AddButton "17", "6", "Sobre o SGCH", 12

'    '# Repaint Ribbon
'    ACPRibbon1.Refresh
    montaTabMenu
    StatusBar1.Panels(1).Width = 1840
    StatusBar1.Panels(2).Width = 4440.189
    StatusBar1.Panels(1).Text = Format(Date, "dd/mm/yyyy")
    StatusBar1.Panels(2).Text = "Usuário: " & NomUsu
    StatusBar1.Panels(3).Text = "Grupo: " & GrupoUsu
    StatusBar1.Panels(4).Text = "DB: " & sServerName & " (" & sDatabaseName & ")"
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
    SqlMenu = "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,1,'TAB','Cadastros','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,1,'CAT','Colaboradores','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,2,'CAT','Candidatos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,3,'CAT','Departamentos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,4,'CAT','Setores','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,5,'CAT','Cargos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,6,'CAT','Habilidades funcionais','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,7,'CAT','Formação escolar','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(1,8,'CAT','Avaliações','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,1,'TAB','Recrutamento','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,1,'CAT','Requisição pessoal','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(2,2,'CAT','Processo seletivo','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,1,'TAB','Capacitação','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,1,'CAT','Cursos/treinamentos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,2,'CAT','Matriz capacitação','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,3,'CAT',' INTD ','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,4,'CAT','Programação','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,5,'CAT','Restrições','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(3,6,'CAT',' ADP ','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,1,'TAB','Configurações','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,1,'CAT','Usuários','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,2,'CAT','Grupos','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,3,'CAT','Sistema','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(4,4,'CAT','PDO','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,1,'TAB','Sobre','" & vCodcoligada & "');" & _
              "Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,1,'CAT','Sobre SGCH','" & vCodcoligada & "');Insert into tbMenu(idmenu,idsub,tipo,nome,codcoligada) Values(5,2,'CAT','Ajuda do SGCH','" & vCodcoligada & "');"
    
    rsMenu.Open SqlMenu, cnBanco
    cnBanco.CommitTrans
    Set rsMenu = Nothing
End Sub

Private Sub Form_Resize()
    '# this procedure will resize the ribbon
    ACPRibbon1.Resize
End Sub

Private Sub ACPRibbon1_MainMenuClick()
    Dim flag As String
    vHelp = WinHelp(HWnd, flag, HELP_QUIT, 0)
    vFechar = 1
    Unload Me
End Sub

Private Sub ACPRibbon1_MenuClick(ByVal ID As String, ByVal Caption As String)
    If ID = 1 Then
        'MsgBox "MenuClick: " & ID & "--" & Caption
        Set chamaForm = New frmRelatorios
        
        chamaForm.Show 1
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If vFechar = 1 Then
        If MsgBox("Deseja Realizar Logout/login", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            vFechar = 0
            Unload frmMenu2
            frmSplash.Show
        Else
            Cancel = 1
        End If
        Exit Sub
    End If
    
    If MsgBox("Deseja encerrar a aplicação", vbQuestion + vbYesNo, "SGCH") = vbYes Then
        cnBanco.Close
        Set cnBanco = Nothing
        End
    End If
    
'    If MsgBox("Deseja encerrar a aplicação", vbQuestion + vbYesNo, "SGCH") = vbYes Then
'        Unload frmMenu2
'        frmSplash.Show
'        Exit Sub
'    End If
    Cancel = 1
    Exit Sub
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
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddTab "1", "Cadastros"
    ACPRibbon1.AddCat "1", "1", "Cadastros gerais", False
    ACPRibbon1.AddCat "2", "1", "Tabelas Auxiliares", True
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "0", "1", "Colaboradores", 2
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "1", "1", "Candidatos", 1
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "2", "1", "Departamentos", 3
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "3", "1", "Setores", 4
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "4", "1", "Cargos", 5
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "5", "2", "Habilidades funcionais", 8
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "6", "2", "Formação escolar", 9
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "7", "2", "Avaliações", 16
    rsConf.MoveNext
    
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddTab "2", "Recrutamento"
    rsConf.MoveNext
    ACPRibbon1.AddCat "3", "2", "Seleção de pessoal", True
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "8", "3", "Requisição de pessoal", 14
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "9", "3", "Processo seletivo", 15
    rsConf.MoveNext
    
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddTab "3", "Capacitação"
    rsConf.MoveNext
    ACPRibbon1.AddCat "4", "3", "Capacitação de pessoal", False
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "10", "4", "Cursos/treinamentos", 6
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "11", "4", "Matriz de capacitação", 7
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "12", "4", " INTD ", 18, , "Identificação das Necessidades de Treinamento e Desenvolvimento"
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "13", "4", "Programação", 13
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "14", "4", "Restrições", 17
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "15", "4", " ADP ", 21, , "Avaliação de Desempenho Profissional"
    rsConf.MoveNext
   
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddTab "4", "Configurações"
    rsConf.MoveNext
    ACPRibbon1.AddCat "5", "4", "Parametrizações", False
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "16", "5", "Usuários", 10
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "17", "5", "Grupos", 11
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "18", "5", "Sistema", 12
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "19", "5", "PDO", 19, , "Processo Decisório Organizacional"
    rsConf.MoveNext
    
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddTab "5", "Sobre"
    rsConf.MoveNext
    ACPRibbon1.AddCat "6", "5", "Ajuda", True
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "20", "6", "Sobre o SGCH", 20
    rsConf.MoveNext
    If rsConf.Fields(5) = "S" Then ACPRibbon1.AddButton "21", "6", "Ajuda do SGCH", 23
    rsConf.MoveNext
    ACPRibbon1.Refresh
    
    Dim rsConfBot As New ADODB.Recordset
    Dim SqlConfBot As String
'    SqlConfBot = "Select * from tbconfgrupo Where codcoligada = '" & vCodcoligada & "' and idmenu = 0 and tbconfgrupo.idgrupo = '" & XCodGrp & "' order by id"
    SqlConfBot = "Select * from tbconfgrupo Where idmenu = 0 and tbconfgrupo.idgrupo = '" & XCodGrp & "' order by id"
    rsConfBot.Open SqlConfBot, cnBanco, adOpenKeyset, adLockReadOnly
    
    vInc = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vEdi = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vSal = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vExc = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vImp = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vFil = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vAva = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vAdi = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vDem = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vAdiRes = rsConfBot.Fields(5)
        rsConfBot.MoveNext
    vAdiRep = rsConfBot.Fields(5)
   
   rsConfBot.Close
    Set rsConfBot = Nothing
End Sub
