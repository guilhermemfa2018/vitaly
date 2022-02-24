VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmimportarnfe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar nfe"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16425
   ControlBox      =   0   'False
   Icon            =   "frmimportarnfe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   16425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   14160
      TabIndex        =   37
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selecione a COLIGADA:"
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
      Left            =   4080
      TabIndex        =   35
      Top             =   240
      Width           =   6255
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
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Gravar Dados"
      Height          =   375
      Left            =   12120
      TabIndex        =   31
      Top             =   7920
      Width           =   1815
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
      Left            =   9840
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Text            =   "CORPORERM_SOBRA"
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
         Text            =   "IMPORT_NFE_TESTE"
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
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "SENHA:"
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
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   7800
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
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16215
      Begin VB.Frame Frame4 
         Caption         =   "Período:"
         Height          =   735
         Left            =   10440
         TabIndex        =   38
         Top             =   120
         Width           =   3735
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2040
            TabIndex        =   40
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   141950977
            CurrentDate     =   43493
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   141950977
            CurrentDate     =   43493
         End
      End
      Begin VB.CommandButton Command1f 
         BackColor       =   &H00808080&
         Caption         =   "Localizar e Importar xml"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   360
         Width           =   3495
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6495
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   11456
         LabelEdit       =   1
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
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   15975
         Begin VB.TextBox Text115 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            TabIndex        =   43
            Top             =   2160
            Width           =   2055
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3240
            TabIndex        =   42
            Text            =   "-"
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
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
            Left            =   6120
            TabIndex        =   17
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   15
            Top             =   1560
            Width           =   5895
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
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
            Left            =   6120
            TabIndex        =   9
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   7440
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   5895
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            Left            =   1800
            TabIndex        =   5
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Contador:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   44
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "CNPJ Coligada:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   41
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Label19"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   6375
         End
         Begin VB.Label Label18 
            Caption         =   "Valor NF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   16
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "CNPJ"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Serie"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   12
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Entrada"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Emissão"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   10
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor / Emissor."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Núm. Chave NFiscal."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   120
            Width           =   4935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Núm. Nota Fiscal."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   2535
         End
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   30
      Top             =   7920
      Width           =   10455
   End
End
Attribute VB_Name = "frmimportarnfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private achei As Boolean
Private Contador As Integer
Private cnpjColigada As String
Private vCodColigada As Integer



'Private Sub Command1_Click()
'    If CommonDialog1.FileName = "" Then
'        MsgBox "É necessario localizar uma nota fiscal de entrada"
'        Exit Sub
'    End If
'    On Error Resume Next
'    Dim DOC As DOMDocument, Temp(3) As String
'    Set DOC = New DOMDocument
'
'    Dim cprod As String, nitem As String, vuncom As String, qcom As String, xprod As String, vprod As String, cEAN As String
'    Dim CFOP As String, NCM As String
'    Dim sBn As String, uCom As String
'    Dim qtdProd As String
'
'   ' Dim qcom
'    Dim vipi As String
'    Dim vicmsst As String
'    Dim vfrete As String
'    Dim vicmsdeson As String
'    Dim pICMS As String
'    Dim vbc As String
'    Dim vICMS As String
'    Dim voutro As String
'    Dim vdesc As String
'    Dim vseg As String
'    Dim nNF As String
'
'    Dim caminho As String
'    Dim XMLdoc As Object
'    Dim i As Integer
'    Dim CountRow As Integer
'    Set XMLdoc = CreateObject("Microsoft.XMLDOM")
'    XMLdoc.async = False
   
'    caminho = CommonDialog1.FileName
'    XMLdoc.Load (caminho)
    
'    Text2.Text = CommonDialog1.FileTitle
'    fA = Text2.Text
'    Text2.Text = ""
'    Text2.Text = Left(fA, InStr(fA, ".") - 1)
'    Text2.Text = Replace(Text2.Text, "_procNFe", "")
'
'
'    Text1.Text = Val(RetornaTagXML((Trim(caminho)), "ide", "cNF")) 'RETORNA O NÚMERO DA NF
'    Text13.Text = (RetornaTagXML((Trim(caminho)), "emit", "xNome"))
'    Text3.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vProd"))
'    Text4.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vDesc"))
'    Text11.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vICMS"))
'    Text12.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vST"))
'    Text10.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vIPI"))
'    Text5.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vFrete"))
'    Text9.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vNF"))
'    Text8.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vOutro"))
'    Text15.Text = (RetornaTagXML((Trim(caminho)), "ide", "dhEmi"))
'    Text15.Text = Left(Text15.Text, InStr(Text15, "T") - 1)
'
'    Text6.Text = Val(RetornaTagXML((Trim(caminho)), "ide", "serie")) 'RETORNA A SERIE DA NOTA
'
'
'    Text14.Text = (RetornaTagXML((Trim(caminho)), "ide", "dhSaiEnt"))
'    Text14.Text = Left(Text14.Text, InStr(Text14, "T") - 1)
'    Text14.Text = Format(Text14.Text, "dd,mm,yyyy")
'    Text15.Text = Format(Text15.Text, "dd,mm,yyyy")
'    If textovalor = "" Then
'        MsgBox "A caixa de calculo porcentagem de venda devera ser preenchida"
'        Exit Sub
'    End If
'    qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").length 'Contando quantos itens tem o nó det (detalhes)
'    For i = 0 To qtdProd - 1 'Varrendo todos os itens
'        cprod = CStr(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text)
'        nitem = CStr(XMLdoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).Value)
'        vuncom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vUnCom").Item(0).Text, ".", ",")
'        qcom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/qCom").Item(0).Text, ".", ",")
'        xprod = CStr(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/xProd").Item(0).Text)
'        vprod = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vProd").Item(0).Text, ".", ",")
'        cEAN = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cEAN").Item(0).Text, ".", ",")
'        CFOP = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/CFOP").Item(0).Text, ".", ",")
'        NCM = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/NCM").Item(0).Text, ".", ",")
'        uCom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/uCom").Item(0).Text, ".", ",")
'        qcom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/qcom").Item(0).Text, ".", ",")
'        vuncom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vuncom").Item(0).Text, ".", ",")
'        'data validade e data fabricação
'        dval = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/med/dVal").Item(0).Text, ".", ",")
'        dFab = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/med/dFab").Item(0).Text, ".", ",")
'        'imposto
'        orig = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/orig").Item(0).Text, ".", ",")
'        CST = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("PIS/PISAliq").Item(0).Text, ".", ",")
'
'
'   Connect
' rs.Open "SELECT * FROM produto", CON, adOpenStatic, adLockOptimistic 'WHERE Código ='" & Text5 & "'"", CON, adOpenStatic, adLockOptimistic"
'
'       rs.AddNew
'       rs!codig = Right(cEAN, 13)
'       rs!CATEGORIA = "Cadastrado Pela entrada NFe" ' Text9.Text
'       rs!descricao = xprod
'       vPorcento = textovalor '3%vValorFinal =
'
'       bb = (((vprod) / 100) * vPorcento + vprod)
'       rs!valor = Format(bb, "##,#0.00")
'       rs!Imagem = "C:\cadasro de clientes\g2o1.jpg"
'       rs!QUANTIDADEMINIMA = 1
'       rs!QUANTIDADEESTOQUE = qcom
'       rs!valorcompra = vprod 'Format(Text15.Text, "##,#0.00")
'       rs!fornc = Text13
'       rs!Grupo = "Não Cadastrado" 'Combo1.Text
'       rs!subgrupo = "Não Cadastrado" 'Combo2.Text
'       rs!desconto = "0,00"
'       rs!clasfiscal = "Não Cadastrado" ' Text31.Text
'       'rs!Info = 'Text34.Text
'       rs!datacadastro = dFab
'       rs!validade = dval
'       'rs!aliquota = Text32
'       rs!orgn = orig
'       rs!CFOP = CFOP
'      ' rs!ipi222 = Text23
'       'rs!pesobruto = Text21
'       'rs!pesoliquido = Text22
'       'rs!icsms = Text24
'       rs!CSOSN = Text25
'       rs!NCM = NCM
'      ' rs!pisconfins = Text27
'       rs!ctspis = CST
'       'rs!aliquota = Text23
'       'rs!REDicsms = Text41
'       'rs!REDicsmsST = Text42
'      ' rs!icsmsST = Text40
'       'rs!CSTICMS = Text37
'       'rs!CSTCONFINS = Text39
'       'rs!CSTIPI = Text38
'       'rs!tribf = Text43
'       'rs!tribe = Text44
'       'rs!tribm = Text45
'       rs!seto = "Não informado"
'       rs!comiss = "0,00"
'  rs.Update
'
'
'      Next i
'    Set XMLdoc = Nothing
'    MsgBox "Cadastrado com Sucesso"
'End Sub

Private Sub Command1f_Click()
    On Error Resume Next
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
        If arquivo.Name Like "*.xml" Then
            caminho = Label19 & "\" & arquivo.Name
            XMLdoc.Load (caminho)
            Text2.Text = arquivo.Name
            fA = Text2.Text
            Text2.Text = ""
            Text2.Text = Left(fA, InStr(fA, ".") - 1)
            vContaXML = vContaXML + 1
            Text115.Text = vContaXML
            If Val(RetornaTagXML((Trim(caminho)), "ide", "mod")) = 55 Then ' NF - Nota Fiscal Eletronica
                Text1.Text = Format(Val(RetornaTagXML((Trim(caminho)), "ide", "nNF")), "000000000") 'RETORNA O NÚMERO DA NF
                Text13.Text = (RetornaTagXML((Trim(caminho)), "emit", "xNome"))
                Text3.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vProd"))
                Text4.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vDesc"))
                Text11.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vICMS"))
                Text12.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vST"))
                Text10.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vIPI"))
                Text5.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vFrete"))
                Text9.Text = (RetornaTagXML((Trim(caminho)), "ICMSTot", "vNF"))
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
                If Text1 <> "" Then
                    IncluirLV ListView1, Text115, Text1, Text6, Text7, Text13, Text15, Text14, Text16, Text2, Text10, Text1, Text1, Text1, Text1, Text1
                End If
            ElseIf Val(RetornaTagXML((Trim(caminho)), "ide", "mod")) = 57 Then ' CT - Conhecimento de Transporte
                If (RetornaTagXML((Trim(caminho)), "toma3", "toma")) = 3 Then
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
                        cnpjColigada = Val(RetornaTagXML((Trim(caminho)), "toma3", "CNPJ")) 'CNPJ DA COLIGADA
                        Text10.Text = Val(RetornaTagXML((Trim(caminho)), "toma3", "CNPJ"))  'CNPJ DA COLIGADA
                        If Text10 = "" Then Text10 = "-"
                        qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").length 'Contando quantos itens tem o nó det (detalhes)
                
                        Text16.Text = Format(Text16.Text / 100, "#,##0.00;(#,##0.00)")
                        If Text1 <> "" Then
                            IncluirLV ListView1, Text115, Text1, Text6, Text7, Text13, Text15, Text14, Text16, Text2, Text10, Text1, Text1, Text1, Text1, Text1
                        End If
                
                    End If
                ElseIf (RetornaTagXML((Trim(caminho)), "toma4", "toma")) = 4 Then
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
                        cnpjColigada = Val(RetornaTagXML((Trim(caminho)), "toma4", "CNPJ")) 'CNPJ DA COLIGADA
                        Text10.Text = Val(RetornaTagXML((Trim(caminho)), "toma4", "CNPJ"))  'CNPJ DA COLIGADA
                        If Text10 = "" Then Text10 = "-"
                        qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").length 'Contando quantos itens tem o nó det (detalhes)
                
                
                        Text16.Text = Format(Text16.Text / 100, "#,##0.00;(#,##0.00)")
                        If Text1 <> "" Then
                            IncluirLV ListView1, Text115, Text1, Text6, Text7, Text13, Text15, Text14, Text16, Text2, Text10, Text1, Text1, Text1, Text1, Text1
                        End If
                    End If
                End If
            End If
        End If
'        Text16.Text = Format(Text16.Text / 100, "#,##0.00;(#,##0.00)")
'        If Text1 <> "" Then
'            IncluirLV ListView1, Text115, Text1, Text6, Text7, Text13, Text15, Text14, Text16, Text2, Text10, Text1, Text1, Text1, Text1, Text1
'        End If
        Contador = Contador + 1
    Next
    
    
    ListView1.Sorted = True
    ListView1.SortKey = 0
    ListView1.SortOrder = lvwAscending
    
    
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
    DTPicker1 = Date
    DTPicker2 = Date
    CriarBancoDeDadosADO
    CriarTabelasADO
    Conexao
SelecionaColigada
    listview_cabecalho
  'textgrid
 
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
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 20
    ListView1.ColumnHeaders.Add , , "NF", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Serie", ListView1.Width / 20
    ListView1.ColumnHeaders.Add , , "CNPJ", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Fornecedor", ListView1.Width / 3.5
    ListView1.ColumnHeaders.Add , , "Emissão", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Entrada", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Valor NF", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Chave NF", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "CNPJ Coligada", ListView1.Width / 5
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub SelecionaColigada()
    CompoeCombo1 Combo1, "corporerm.dbo.GCOLIGADA", "codcoligada", "nomefantasia"
End Sub

Private Sub adivinhaColigada(cnpjDaColigada As String)
    Dim rsAchaColigada As New ADODB.Recordset
    Dim SqlAchaColigada As String
    SqlAchaColigada = "select * from corporerm.dbo.GCOLIGADA where CGC = '" & Format(cnpjDaColigada, "00\.000\.000\/0000\-00") & "'"
    rsAchaColigada.Open SqlAchaColigada, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaColigada.RecordCount > 0 Then
        'Combo1.Text = Format(rsAchaColigada.Fields(0), "000000") & "-" & rsAchaColigada.Fields(1)
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
            SqlSalvar = "Select * from tbNFE where nfe = " & Val(ListView1.SelectedItem.ListSubItems.Item(1)) & " and serie = '" & ListView1.SelectedItem.ListSubItems.Item(2) & "'"
            rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
            If rsSalvar.RecordCount = 0 Then
                rsSalvar.AddNew
                rsSalvar.Fields(1) = ListView1.SelectedItem.ListSubItems.Item(1) 'ListView1.ListItems.Item(X) ' NFE
                rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2) ' SERIE
                rsSalvar.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(3) 'CNPJ
                rsSalvar.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(4) 'FORNECEDOR
                If ListView1.SelectedItem.ListSubItems.Item(5) <> "" Then
                    rsSalvar.Fields(5) = ListView1.SelectedItem.ListSubItems.Item(5) 'DATA EMISSAO
                End If
                
                If ListView1.SelectedItem.ListSubItems.Item(6) <> "" Then
                    rsSalvar.Fields(6) = ListView1.SelectedItem.ListSubItems.Item(6) 'DATA ENTRADA
                End If
                If ListView1.SelectedItem.ListSubItems.Item(7) = "" Then
                    rsSalvar.Fields(7) = 0 'VALOR NF
                Else
                    rsSalvar.Fields(7) = ListView1.SelectedItem.ListSubItems.Item(7) 'VALOR NF
                End If
                rsSalvar.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(8) 'CHAVE NF
                rsSalvar.Fields(9) = Date   'DATA DE CADASTRO
                
'                    rsSalvar.Fields(10) = Val(Mid$(Combo1.Text, 1, 6)) 'CODIGO DA COLIGADA
                
                If ListView1.SelectedItem.ListSubItems.Item(9) <> "-" Then
                    adivinhaColigada ListView1.SelectedItem.ListSubItems.Item(9)
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


'LIXO




'            FlexGrid.Clear
 
'            For i = 0 To qtdProd - 1 'Varrendo todos os itens
'                cprod = CStr(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cProd").Item(0).Text)
'                nitem = CStr(XMLdoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).Value)
'                vuncom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vUnCom").Item(0).Text, ".", ",")
'                qcom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/qCom").Item(0).Text, ".", ",")
'                xprod = CStr(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/xProd").Item(0).Text)
'                vprod = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vProd").Item(0).Text, ".", ",")
'                cEAN = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/cEAN").Item(0).Text, ".", ",")
'                CFOP = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/CFOP").Item(0).Text, ".", ",")
'                NCM = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/NCM").Item(0).Text, ".", ",")
'                uCom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/uCom").Item(0).Text, ".", ",")
'                qcom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/qcom").Item(0).Text, ".", ",")
'                vuncom = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vuncom").Item(0).Text, ".", ",")
'                'data validade e data fabricação
'                dval = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/med/dVal").Item(0).Text, ".", ",")
'                dFab = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/med/dFab").Item(0).Text, ".", ",")
'                'imposto
'                orig = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/ICMS/ICMS60/orig").Item(0).Text, ".", ",")
'
'
'                vdesc = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("prod/vDesc").Item(0).Text, ".", ",")
 '               CST = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/CST").Item(0).Text, ".", ",")
 '               vbc = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/vBC").Item(0).Text, ".", ",")
 '               vCOFINS = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/COFINS/COFINSAliq/vCOFINS").Item(0).Text, ".", ",")
 '               vICMSSTRet = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/ICMS/ICMS60/vICMSSTRet").Item(0).Text, ".", ",")
 '               vPIS = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("imposto/PIS/PISAliq/vPIS").Item(0).Text, ".", ",")
 '               infAdProd = Replace(XMLdoc.selectNodes("nfeProc/NFe/infNFe/det").Item(i).selectNodes("det/infAdProd").Item(0).Text, ".", ",")
'
'
'
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 1) = cprod
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 2) = xprod
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 3) = qcom ', DT_RIGHT 'Unid
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 4) = uCom ', DT_RIGHT 'Quant
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 5) = Format(vuncom, "###,##0.00") ', DT_RIGHT 'Valor Unitário
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 6) = Format(vprod, "###,##0.00") ', DT_RIGHT
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 7) = nitem ', DT_CENTER
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 8) = cEAN ', DT_CENTER
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 9) = CFOP ', DT_CENTER
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 10) = NCM ', DT_LEFT
'
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 11) = vdesc
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 12) = CST
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 13) = vbc
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 14) = vCOFINS
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 15) = vICMSSTRet
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 16) = vPIS
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 17) = orig
'                FlexGrid.TextMatrix(FlexGrid.Rows - 1, 18) = Format(dval, "dd/mm/yyyy")
'                'FlexGrid.TextMatrix(FlexGrid.rows - 1, 19) = vdesc
'                'FlexGrid.TextMatrix(FlexGrid.rows - 1, 20) = nNF
'                FlexGrid.Rows = FlexGrid.Rows + 1
'            Next i



