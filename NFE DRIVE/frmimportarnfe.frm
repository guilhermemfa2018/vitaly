VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmimportarnfe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar nfe"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmimportarnfe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   18600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   9360
      TabIndex        =   52
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   10320
      TabIndex        =   51
      Text            =   "voltar propriedade viseble = true para exibir a chave da NF e poder copia-la"
      Top             =   9600
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Localizar e Importa XMLs para a tabela de controle de NFe"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   39
      Top             =   360
      Width           =   7095
      Begin VB.Frame Frame9 
         Caption         =   "Não Importados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   49
         Top             =   240
         Width           =   1695
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Analisados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         TabIndex        =   47
         Top             =   240
         Width           =   1575
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1f 
         BackColor       =   &H80000004&
         Caption         =   "Localizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Importar"
         Height          =   495
         Left            =   5400
         TabIndex        =   40
         ToolTipText     =   "Importa dados listados acima para a tabela de controle de NFe"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione os dados para impessão do relatório"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   37
      Top             =   360
      Width           =   9975
      Begin VB.Frame Frame6 
         Caption         =   "Coligada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   44
         Top             =   240
         Width           =   4695
         Begin VB.ComboBox Combo1 
            Height          =   345
            Left            =   120
            TabIndex        =   45
            Text            =   "000005-VITALY INDUSTRIA MECANICA EIRELI - ME"
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   3495
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1920
            TabIndex        =   42
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   140247041
            CurrentDate     =   43493
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   140247041
            CurrentDate     =   43493
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   8640
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Configuração de conexão DB RM Sistemas "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   12120
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Text            =   "CORPORERM"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   25
         Text            =   "vigamax"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Text            =   "sa"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   23
         Text            =   "IMPORT_NFE"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "SRV1002\CORPORERM"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Banco RM:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "SENHA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "USUÁRIO:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Leitura e Importação das NFEs"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18375
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   7680
         TabIndex        =   1
         Top             =   4800
         Visible         =   0   'False
         Width           =   10455
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7440
            TabIndex        =   63
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   61
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            TabIndex        =   60
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   59
            Top             =   2880
            Width           =   2775
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   58
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox Text115 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   35
            Top             =   2160
            Width           =   2055
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   34
            Text            =   "-"
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            TabIndex        =   17
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   1560
            Width           =   5895
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8640
            TabIndex        =   13
            Top             =   960
            Width           =   855
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   9600
            Top             =   2040
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            TabIndex        =   9
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7440
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Produto"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "ICMS"
            Height          =   255
            Left            =   7440
            TabIndex        =   62
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "ST"
            Height          =   255
            Left            =   5760
            TabIndex        =   57
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Frete"
            Height          =   255
            Left            =   4080
            TabIndex        =   56
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   1200
            TabIndex        =   55
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Mod"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Contador:"
            Height          =   255
            Left            =   6240
            TabIndex        =   36
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "CNPJ Coligada:"
            Height          =   255
            Left            =   3240
            TabIndex        =   33
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Caminho do diretorio"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   3840
            Width           =   9975
         End
         Begin VB.Label Label18 
            Caption         =   "Valor NF"
            Height          =   255
            Left            =   6120
            TabIndex        =   16
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "CNPJ"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Serie"
            Height          =   255
            Left            =   8640
            TabIndex        =   12
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Entrada"
            Height          =   375
            Left            =   7440
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Emissão"
            Height          =   375
            Left            =   6120
            TabIndex        =   10
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor / Emissor."
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Núm. Chave NFiscal."
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   120
            Width           =   4935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Núm. Nota Fiscal."
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7815
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   18135
         _ExtentX        =   31988
         _ExtentY        =   13785
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   30
      Top             =   9720
      Width           =   8535
   End
End
Attribute VB_Name = "frmimportarnfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private achei As Boolean
Private Contador As Integer
Private Contador2 As Integer
Private cnpjColigada As String
Private vCodColigada As Integer

Private Sub Command1f_Click()
    On Error Resume Next
    Label5.Caption = "Aguarde! Analisando arquivos XML"
    Dim DOC As DOMDocument, Temp(3) As String
    Set DOC = New DOMDocument
   
    Dim cprod As String, nitem As String, vuncom As String, qcom As String, xprod As String, vprod As String, cEAN As String
    Dim CFOP As String, NCM As String
    Dim sBn As String, uCom As String
    Dim qtdProd As String
       
   ' Dim qcom
    Dim vipi As String
    Dim vicmsst As String
    Dim vfrete As String
    Dim vicmsdeson As String
    Dim pICMS As String
    Dim vbc As String
    Dim vICMS As String
    Dim voutro As String
    Dim vdesc As String
    Dim vseg As String
    Dim nNF As String
    
    Dim caminho As String
    Dim XMLdoc As Object
    Dim i As Integer
    Dim CountRow As Integer
    Contador = 0
    Contador2 = 0
    
    Set XMLdoc = CreateObject("Microsoft.XMLDOM")
    XMLdoc.async = False
    CommonDialog1.Filter = "Directories|*.XML" 'set files-filter to show dirs only
    CommonDialog1.ShowOpen
   
   
'    caminho = CommonDialog1.FileName
'    XMLdoc.Load (caminho)
    
    '********************************************************************
    'PEGA O CAMINHO DO DIRETORIO SELECIONADO
    Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir    'Remember the current active directory
    CommonDialog1.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog1.CancelError = True 'allow escape key/cancel

    If Err <> 32755 Then    ' User didn't chose Cancel.
        Label19.Caption = CurDir
    End If
    ChDir sTempDir  'restore path to what it was at entering
    'GUARDA O CAMINHO DO DIRETORIO SELECIONADO
    Label19 = CurDir
    '********************************************************************
    
    '********************************************************************
    'LISTA TODOS OS ARQUIVOS XML DO DIRETÓRIO SELECIONADO
    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim diretorio As Folder
    Set diretorio = fso.GetFolder(Label19.Caption)
    '********************************************************************
    
    Dim arquivo As File
    Dim subdiretorio As Folder
    Dim vContaXML As Integer
    vContaXML = 0
    For Each arquivo In diretorio.Files
        
        If Mid(arquivo.Name, 21, 2) = 55 Or Mid(arquivo.Name, 21, 2) = 57 Then
            If localizaNF(Mid(arquivo.Name, 1, Len(arquivo.Name) - 4)) = True Then
                If arquivo.Name Like "*.xml" Then
                    caminho = Label19 & "\" & arquivo.Name
                    XMLdoc.Load (caminho)
                    Text2.Text = arquivo.Name
                    fA = Text2.Text
                    Text2.Text = ""
                    Text2.Text = Left(fA, InStr(fA, ".") - 1)
                    Text2.Text = Mid$(Text2.Text, 1, 44)
                    
                    vContaXML = vContaXML + 1
                    Text115.Text = vContaXML
                    If Val(RetornaTagXML((Trim(caminho)), "ide", "mod")) = 55 Then ' NF - Nota Fiscal Eletronica
                        Text1.Text = Format(Val(RetornaTagXML((Trim(caminho)), "ide", "nNF")), "000000000") 'RETORNA O NÚMERO DA NF
                        Text13.Text = (RetornaTagXML((Trim(caminho)), "emit", "xNome"))
                        Text20.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vProd"))
                        Text21.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vDesc"))
                        Text24.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vICMS"))
                        Text23.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vST"))
                        Text10.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vIPI"))
                        Text22.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vFrete"))
                        Text8.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vOutro"))
                        Text15.Text = (RetornaTagXML((Trim(caminho)), "ide", "dhEmi"))
                        Text15.Text = Left(Text15.Text, InStr(Text15, "T") - 1)
                        Text14.Text = (RetornaTagXML((Trim(caminho)), "ide", "dhSaiEnt"))
                        Text14.Text = Left(Text14.Text, InStr(Text14, "T") - 1)
                        Text14.Text = Format(Text14.Text, "dd/mm/yyyy")
                        Text15.Text = Format(Text15.Text, "dd/mm/yyyy")
                        Text6.Text = Val(RetornaTagXML((Trim(caminho)), "ide", "serie")) 'RETORNA A SERIE DA NOTA
                        Text7.Text = Val(RetornaTagXML((Trim(caminho)), "emit", "CNPJ")) 'CNPJ DO EMISSOR
                        Text16.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vNF")) ' TOTAL DA NF
                        cnpjColigada = Val(RetornaTagXML((Trim(caminho)), "dest", "CNPJ")) 'CNPJ DA COLIGADA
                        Text10.Text = Val(RetornaTagXML((Trim(caminho)), "dest", "CNPJ"))  'CNPJ DA COLIGADA
                        
                        qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").length 'Contando quantos itens tem o nó det (detalhes)
                        
                        Text16.Text = Format(Text16.Text / 100, "#,##0.00;(#,##0.00)")
                        
                        Text19.Text = (RetornaTagXML((Trim(caminho)), "ide", "mod")) + " - NFe"
                        
                        If Text1 <> "" Then
                            IncluirLV ListView1, Text1, Text6, Text7, Text13, Text15, Text14, Text16, Text2, Text10, Text19, Text1, Text1, Text1, Text1, Text1
                            ListView1.Refresh
                            Contador2 = Contador2 + 1
                            Debug.Print Str(Contador) + " - NF"
                        End If
                    ElseIf Val(RetornaTagXML((Trim(caminho)), "ide", "mod")) = 57 And Val(RetornaTagXML((Trim(caminho)), "ide", "toma")) > 0 Then ' CT - Conhecimento de Transporte
                        
                        
                        If Val(RetornaTagXML((Trim(caminho)), "dest", "CNPJ")) = "24874889000102" Or Val(RetornaTagXML((Trim(caminho)), "dest", "CNPJ")) = "19431980000105" Then
                                Text1.Text = Format(Val(RetornaTagXML((Trim(caminho)), "ide", "nCT")), "000000000") 'RETORNA O NÚMERO DO CT
                                Text13.Text = (RetornaTagXML((Trim(caminho)), "emit", "xNome"))
                                Text15.Text = (RetornaTagXML((Trim(caminho)), "ide", "dhEmi"))
                                Text15.Text = Left(Text15.Text, InStr(Text15, "T") - 1)
                                Text14.Text = (RetornaTagXML((Trim(caminho)), "ide", "dhSaiEnt"))
                                Text14.Text = Left(Text14.Text, InStr(Text14, "T") - 1)
                                Text14.Text = Format(Text14.Text, "dd/mm/yyyy")
                                Text15.Text = Format(Text15.Text, "dd/mm/yyyy")
                                Text6.Text = Val(RetornaTagXML((Trim(caminho)), "ide", "serie")) 'RETORNA A SERIE DA NOTA
                                Text7.Text = Val(RetornaTagXML((Trim(caminho)), "emit", "CNPJ")) 'CNPJ DO EMISSOR
                                Text16.Text = (RetornaTagXML((Trim(caminho)), "vPrest", "vTPrest")) ' TOTAL DA NF
                        
                                cnpjColigada = Val(RetornaTagXML((Trim(caminho)), "dest", "CNPJ")) 'CNPJ DA COLIGADA
                                Text10.Text = Val(RetornaTagXML((Trim(caminho)), "dest", "CNPJ"))  'CNPJ DA COLIGADA
                        
                        
                                If Text10 = "" Then Text10 = "-"
                                qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").length 'Contando quantos itens tem o nó det (detalhes)
                        
                                Text16.Text = Format(Text16.Text / 100, "#,##0.00;(#,##0.00)")
                                
                                If (RetornaTagXML((Trim(caminho)), "toma3", "toma")) = 3 Then
                                    Text19.Text = (RetornaTagXML((Trim(caminho)), "ide", "mod")) + " - CTe(toma3)"
                                ElseIf (RetornaTagXML((Trim(caminho)), "toma4", "toma")) = 4 Then
                                    Text19.Text = (RetornaTagXML((Trim(caminho)), "ide", "mod")) + " - CTe(toma4)"
                                End If
                                
                                If Text1 <> "" Then
                                    IncluirLV ListView1, Text1, Text6, Text7, Text13, Text15, Text14, Text16, Text2, Text10, Text19, Text1, Text1, Text1, Text1, Text1
                                    ListView1.Refresh
                                    Contador2 = Contador2 + 1
                                    Debug.Print Str(Contador) + " - " + Text19.Text
                                End If
                        End If
                    End If
                End If
            End If
        End If
        Contador = Contador + 1
        Debug.Print Contador
        Text11.Text = Contador
        Text11.Refresh
        Text12.Text = Contador2
        Text12.Refresh
    Next
    ListView1.Sorted = True
    ListView1.SortKey = 0
    ListView1.SortOrder = lvwAscending
    Label5.Caption = "Análise concluida"
    MsgBox "Análise concluida", vbInformation, "NFRM"
    Set XMLdoc = Nothing
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    GravarDados
End Sub

Private Sub Command4_Click()
    vDataFilter1 = DTPicker1.Value
    vDataFilter2 = DTPicker2.Value
    FCRListaNFE.Show 1
End Sub

Private Sub Form_Load()
    vBancoTotvs = Text9.Text
    DTPicker1 = Date
    DTPicker2 = Date
    CriarBancoDeDadosADO
    CriarTabelasADO
    Conexao
    SelecionaColigada
    listview_cabecalho
End Sub

Public Function RetornaTagXML(strCaminhoXML As String, TagMae As String, SubTag As String) As String
On Error Resume Next
RetornaTagXML = ""
Set xml = New DOMDocument
xml.async = False
If xml.Load(strCaminhoXML) Then
    ' *** Tentar pegar o strCampoXML
    Set objNodeList = xml.getElementsByTagName(TagMae & "//" & SubTag)
    Set objNode = objNodeList.nextNode
    Dim sLeitura As String
    sLeitura = objNode.Text
    If Len(Trim(sLeitura)) > 0 Then 'CONSEGUI LER O XML NODE
        RetornaTagXML = sLeitura
    End If
    Else
    MsgBox "Não foi possível abrir o arquivo XML da NFe especificada para Leitura.", vbCritical, "Erro."
End If
End Function
Private Function GiveMeTheEnd(sIN As String) As String
Dim arrAll() As String
arrAll = Split(sIN, "\")
GiveMeTheEnd = arrAll(UBound(arrAll))
End Function

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "NF", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Serie", ListView1.Width / 23
    ListView1.ColumnHeaders.Add , , "CNPJ", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Fornecedor", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Emissão", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Entrada", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Valor NF", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Chave NF", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "CNPJ Coligada", ListView1.Width / 13
    ListView1.ColumnHeaders.Add , , "Mod", ListView1.Width / 13
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub SelecionaColigada()
    CompoeCombo1 Combo1, "" & vBancoTotvs & ".dbo.GCOLIGADA", "codcoligada", "nomefantasia"
End Sub

Private Sub adivinhaColigada(cnpjDaColigada As String)
    Dim rsAchaColigada As New ADODB.Recordset
    Dim SqlAchaColigada As String
    SqlAchaColigada = "select * from " & vBancoTotvs & ".dbo.GCOLIGADA where CGC = '" & Format(cnpjDaColigada, "00\.000\.000\/0000\-00") & "'"
    rsAchaColigada.Open SqlAchaColigada, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaColigada.RecordCount > 0 Then
        vCodColigada = rsAchaColigada.Fields(0)
    Else
        vCodColigada = 0
    End If
    rsAchaColigada.Close
End Sub


Private Function GravarDados()
    Dim X As Integer
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    '>>>>>> GRAVAR NFS <<<<<<<<<
    If ListView1.ListItems.Count > 0 Then
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            SqlSalvar = "Select * from tbNFE where chavenf = '" & ListView1.SelectedItem.ListSubItems.Item(8) & "'"
            rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
            If rsSalvar.RecordCount = 0 Then
                rsSalvar.AddNew
                rsSalvar.Fields(1) = ListView1.ListItems.Item(X) ' NFE
                rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(1) ' SERIE
                rsSalvar.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2) 'CNPJ
                rsSalvar.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(3) 'FORNECEDOR
                If ListView1.SelectedItem.ListSubItems.Item(4) <> "" Then
                    rsSalvar.Fields(5) = ListView1.SelectedItem.ListSubItems.Item(4) 'DATA EMISSAO
                End If
                
                If ListView1.SelectedItem.ListSubItems.Item(5) <> "" Then
                    rsSalvar.Fields(6) = ListView1.SelectedItem.ListSubItems.Item(5) 'DATA ENTRADA
                End If
                If ListView1.SelectedItem.ListSubItems.Item(6) = "" Then
                    rsSalvar.Fields(7) = 0 'VALOR NF
                Else
                    rsSalvar.Fields(7) = ListView1.SelectedItem.ListSubItems.Item(6) 'VALOR NF
                End If
                rsSalvar.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(7) 'CHAVE NF
                rsSalvar.Fields(9) = Date   'DATA DE CADASTRO
                
                If ListView1.SelectedItem.ListSubItems.Item(8) <> "-" Then
                    adivinhaColigada ListView1.SelectedItem.ListSubItems.Item(8)
                    rsSalvar.Fields(10) = vCodColigada 'CODIGO DA COLIGADA
                End If
            End If
            rsSalvar.Update
            rsSalvar.Close
        Next
        Set rsSalvar = Nothing
    End If
    Label5.ForeColor = &H8000&
    Label5.Caption = "Dados gravados com sucesso"
    Exit Function
TrataErro:
    GravarDados = False
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Function
TrataErro1:
    Resume Next
End Function

Private Function localizaNF(vChaveNF As String)
    localizaNF = True
    Dim rslocalizaNF As New ADODB.Recordset
    Dim SqllocalizaNF As String

    
    SqllocalizaNF = "Select * from tbNFE where chavenf = '" & Mid$(vChaveNF, 1, 44) & "'"
    rslocalizaNF.Open SqllocalizaNF, cnBanco, adOpenKeyset, adLockReadOnly
    If rslocalizaNF.RecordCount > 0 Then
        localizaNF = False
    End If
    rslocalizaNF.Close
End Function

Private Sub ListView1_Click()
'    AlteraLV ListView1, Text17, Text17, Text17, Text17, Text17, Text17, Text17, Text17, Text17, Text18, Text18, Text18, Text18, Text18, Text18
End Sub
