VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Begin VB.Form frmEditor 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option4 
      Height          =   615
      Left            =   2160
      Picture         =   "frmEditor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   7080
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      Height          =   615
      Left            =   1560
      Picture         =   "frmEditor.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   10245
      TabIndex        =   24
      Top             =   7080
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmEditor.frx":1994
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmEditor.frx":19B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmEditor.frx":19CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Instruções"
         Height          =   6015
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   11295
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
            Height          =   5655
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   23
            Text            =   "frmEditor.frx":19E8
            Top             =   240
            Width           =   11055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Texto "
         Height          =   5535
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   11295
         Begin VB.TextBox Text5 
            Height          =   5175
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   11055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Gerais "
         Height          =   6015
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   11295
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Top             =   5520
            Width           =   1815
         End
         Begin VB.Frame Frame4 
            Caption         =   "Configurações "
            Height          =   5655
            Left            =   4080
            TabIndex        =   29
            Top             =   240
            Width           =   7095
            Begin VB.Frame Frame11 
               Caption         =   "Fonte Certificadora "
               Height          =   1095
               Left            =   3600
               TabIndex        =   46
               Top             =   4440
               Width           =   3375
               Begin VB.Frame Frame17 
                  Caption         =   "Tamanho"
                  Height          =   735
                  Left            =   2280
                  TabIndex        =   57
                  Top             =   240
                  Width           =   975
                  Begin VB.ComboBox Combo8 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   58
                     Top             =   240
                     Width           =   735
                  End
               End
               Begin VB.Frame Frame16 
                  Caption         =   "Nome"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   2055
                  Begin VB.ComboBox Combo7 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   56
                     Top             =   240
                     Width           =   1815
                  End
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "Fonte Rodapé"
               Height          =   1095
               Left            =   120
               TabIndex        =   45
               Top             =   4440
               Width           =   3375
               Begin VB.Frame Frame15 
                  Caption         =   "Tamanho"
                  Height          =   735
                  Left            =   2280
                  TabIndex        =   53
                  Top             =   240
                  Width           =   975
                  Begin VB.ComboBox Combo6 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   54
                     Top             =   240
                     Width           =   735
                  End
               End
               Begin VB.Frame Frame14 
                  Caption         =   "Nome"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   51
                  Top             =   240
                  Width           =   2055
                  Begin VB.ComboBox Combo5 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   52
                     Top             =   240
                     Width           =   1815
                  End
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Fonte Corpo"
               Height          =   1095
               Left            =   3600
               TabIndex        =   44
               Top             =   3240
               Width           =   3375
               Begin VB.Frame Frame13 
                  Caption         =   "Tamanho"
                  Height          =   735
                  Left            =   2280
                  TabIndex        =   49
                  Top             =   240
                  Width           =   975
                  Begin VB.ComboBox Combo4 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   50
                     Top             =   240
                     Width           =   735
                  End
               End
               Begin VB.Frame Frame12 
                  Caption         =   "Nome"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   47
                  Top             =   240
                  Width           =   2055
                  Begin VB.ComboBox Combo3 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   48
                     Top             =   240
                     Width           =   1815
                  End
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Fonte Cabeçalho"
               Height          =   1095
               Left            =   120
               TabIndex        =   39
               Top             =   3240
               Width           =   3375
               Begin VB.Frame Frame8 
                  Caption         =   "Tamanho"
                  Height          =   735
                  Left            =   2280
                  TabIndex        =   42
                  Top             =   240
                  Width           =   975
                  Begin VB.ComboBox Combo2 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   43
                     Top             =   240
                     Width           =   735
                  End
               End
               Begin VB.Frame Frame7 
                  Caption         =   "Nome"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   40
                  Top             =   240
                  Width           =   2055
                  Begin VB.ComboBox Combo1 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   41
                     Top             =   240
                     Width           =   1815
                  End
               End
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Fundo"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   1200
               Width           =   855
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Logo"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   840
               Width           =   855
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Borda"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   480
               Width           =   975
            End
            Begin VB.Frame Frame6 
               Caption         =   "Fundo "
               Height          =   2415
               Index           =   0
               Left            =   4680
               TabIndex        =   30
               Top             =   120
               Width           =   2295
               Begin VB.CommandButton Command3 
                  Height          =   495
                  Left            =   600
                  TabIndex        =   34
                  Top             =   1800
                  Width           =   495
               End
               Begin VB.CommandButton Command2 
                  Height          =   495
                  Left            =   120
                  TabIndex        =   33
                  Top             =   1800
                  Width           =   495
               End
               Begin MSComDlg.CommonDialog cdlFoto 
                  Left            =   1680
                  Top             =   1800
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.PictureBox Picture2 
                  Height          =   1455
                  Left            =   120
                  ScaleHeight     =   1395
                  ScaleWidth      =   1995
                  TabIndex        =   31
                  Top             =   240
                  Width           =   2055
                  Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
                     Height          =   1455
                     Left            =   0
                     Top             =   0
                     Width           =   2055
                     _ExtentX        =   3625
                     _ExtentY        =   2566
                     Image           =   "frmEditor.frx":1FB4
                  End
                  Begin VB.Label Label59 
                     Alignment       =   2  'Center
                     Caption         =   "A Imagem não se encontra no local especificado"
                     Height          =   615
                     Left            =   240
                     TabIndex        =   32
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1335
                  End
               End
            End
            Begin VB.Label Label53 
               BackColor       =   &H8000000C&
               Height          =   255
               Left            =   4680
               TabIndex        =   35
               Top             =   2520
               Visible         =   0   'False
               Width           =   2295
            End
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Top             =   3360
            Width           =   3855
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   2760
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   285
            Left            =   2760
            TabIndex        =   22
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   57278465
            CurrentDate     =   40742
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   57278465
            CurrentDate     =   40742
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   4800
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   57278465
            CurrentDate     =   40742
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   3855
         End
         Begin VB.Label Label11 
            Caption         =   "Nota:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   5280
            Width           =   1935
         End
         Begin VB.Label Label10 
            Caption         =   "Carga horária:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Nome do treinamento:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "Data Término:"
            Height          =   255
            Left            =   2760
            TabIndex        =   20
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Data Início:"
            Height          =   255
            Left            =   1440
            TabIndex        =   19
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Título:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Responsável:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   4560
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Data de emissão:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Nome Empresa certificadora:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Participante:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   1095
         End
      End
   End
   Begin VB.OptionButton Option2 
      Height          =   615
      Left            =   960
      Picture         =   "frmEditor.frx":1FCC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Height          =   615
      Left            =   360
      Picture         =   "frmEditor.frx":2C96
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Conexão estabeleciada"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public cnBanco As ADODB.Connection
'Public ADOXCat As New ADOX.Catalog
'Public oConn As ADODB.Connection

Public rsEdText As New ADODB.Recordset
Public sqlEdText As String
Public Caminho1 As String

Private Sub Command1_Click()
    sqlEdText = "Select * from tbTeste"
    rsEdText.Open sqlEdText, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsEdText.RecordCount = 0 Then rsEdText.AddNew
    rsEdText.Fields(1) = Text5.Text
    rsEdText.Fields(2) = Text2.Text
    rsEdText.Fields(3) = Text1.Text
    rsEdText.Fields(4) = DTPicker1.Value
    rsEdText.Fields(5) = DTPicker2.Value
    rsEdText.Fields(6) = DTPicker3.Value
    rsEdText.Fields(7) = Text3.Text
    rsEdText.Fields(8) = Text4.Text
    rsEdText.Fields(9) = Text7.Text
    rsEdText.Fields(10) = Text8.Text
    rsEdText.Fields(14) = Label53.Caption
    If Check1.Value = 1 Then rsEdText.Fields(12) = "S" Else rsEdText.Fields(12) = "N"
    If Check3.Value = 1 Then rsEdText.Fields(13) = "S" Else rsEdText.Fields(13) = "N"
    Text5 = rsEdText.Fields(1)
    rsEdText.Fields(15) = Combo1.Text
    rsEdText.Fields(16) = Combo3.Text
    rsEdText.Fields(17) = Combo5.Text
    rsEdText.Fields(18) = Combo7.Text
    rsEdText.Fields(24) = Text9.Text
    
    'Fonte
    rsEdText.Fields(15) = Combo1.Text
    rsEdText.Fields(16) = Combo3.Text
    rsEdText.Fields(17) = Combo5.Text
    rsEdText.Fields(18) = Combo7.Text
    'Tamanho Fonte
    rsEdText.Fields(19) = Combo2.Text
    rsEdText.Fields(20) = Combo4.Text
    rsEdText.Fields(21) = Combo6.Text
    rsEdText.Fields(22) = Combo8.Text
    'Alinhamento da fonte do Corpo
    rsEdText.Fields(23) = 0
    If Option1.Value = True Then rsEdText.Fields(23) = 1
    If Option2.Value = True Then rsEdText.Fields(23) = 2
    If Option3.Value = True Then rsEdText.Fields(23) = 3
    If Option4.Value = True Then rsEdText.Fields(23) = 4

    MsgBox "Registro gravado com sucesso !", vbInformation, "Gravado"
    rsEdText.Update
    rsEdText.Close
    Set rsEdText = Nothing
    'restauraDados
    FCRCertificado.Show
End Sub

Private Sub Command2_Click()
    'carregar imagem para o Picture
    With cdlFoto
        .Filter = "(Arquivo *.JPG)|*.jpg"
        .ShowOpen
        Caminho1 = .FileName
    End With
    'mostra a figura
    'Image1.Picture = LoadPicture(Caminho1)
    aicAlphaImage1.LoadImage_FromFile (Caminho1)
    Label53 = Caminho1
End Sub

Private Sub Form_Load()
    Conectar
    preencheComboFontes
    preencheComboTamanhoFontes
    restauraDados
End Sub

Private Sub Conectar()
On Error GoTo Erro
    Set cnBanco = New ADODB.Connection
    cnBanco.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=TesteText;Data Source=ET-0001"
    ADOXCat.ActiveConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=TesteText;Data Source=ET-0001"
    
    Label6.Visible = True
    Exit Sub
Erro:
    Label6.ForeColor = &HC0&
    Label6.Visible = True
    Label6.Caption = "Conexão não estabelecida"
End Sub

Private Sub restauraDados()
    sqlEdText = "Select * from tbTeste"
    rsEdText.Open sqlEdText, cnBanco, adOpenKeyset, adLockOptimistic
    
    Text5.Text = rsEdText.Fields(1)
    Text2.Text = rsEdText.Fields(2)
    Text1.Text = rsEdText.Fields(3)
    DTPicker1.Value = rsEdText.Fields(4)
    DTPicker2.Value = rsEdText.Fields(5)
    DTPicker3.Value = rsEdText.Fields(6)
    Text3.Text = rsEdText.Fields(7)
    Text4.Text = rsEdText.Fields(8)
    Text7.Text = rsEdText.Fields(9)
    Text8.Text = rsEdText.Fields(10)
    If Not IsNull(rsEdText.Fields(24)) Then Text9.Text = rsEdText.Fields(24)

    If Not IsNull(rsEdText.Fields(12)) And rsEdText.Fields(12) = "S" Then Check1.Value = 1 Else Check1.Value = 0
    If Not IsNull(rsEdText.Fields(13)) And rsEdText.Fields(13) = "S" Then Check3.Value = 1 Else Check3.Value = 0
    If rsEdText.Fields(14) <> "Null" Then
        On Error GoTo TrataErro1
        Label53.Caption = rsEdText.Fields(14)
        aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    End If
    'Fonte
    Combo1.Text = rsEdText.Fields(15)
    Combo3.Text = rsEdText.Fields(16)
    Combo5.Text = rsEdText.Fields(17)
    Combo7.Text = rsEdText.Fields(18)
    'Tamanho Fonte
    Combo2.Text = rsEdText.Fields(19)
    Combo4.Text = rsEdText.Fields(20)
    Combo6.Text = rsEdText.Fields(21)
    Combo8.Text = rsEdText.Fields(22)
    'Alinhamento Fonte Corpo
    If rsEdText.Fields(23) = 1 Then Option1.Value = True
    If rsEdText.Fields(23) = 2 Then Option2.Value = True
    If rsEdText.Fields(23) = 3 Then Option3.Value = True
    If rsEdText.Fields(23) = 4 Then Option4.Value = True
    
    rsEdText.Close
    Set rsEdText = Nothing
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Sub preencheComboFontes()
    'preenche a combo box com as fontes disponíveis
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(i)
        Combo3.AddItem Screen.Fonts(i)
        Combo5.AddItem Screen.Fonts(i)
        Combo7.AddItem Screen.Fonts(i)
    Next i
    Combo1.Text = "Arial"
    Combo3.Text = "Arial"
    Combo5.Text = "Arial"
    Combo7.Text = "Arial"
End Sub

Private Sub preencheComboTamanhoFontes()
    'preenche a combo box com os tamanhos das fontes
    Dim i As Integer
    For i = 8 To 24 Step 2
        Combo2.AddItem i
        Combo4.AddItem i
        Combo6.AddItem i
        Combo8.AddItem i
    Next i
    Combo2.ListIndex = 0
    Combo4.ListIndex = 0
    Combo6.ListIndex = 0
    Combo8.ListIndex = 0
End Sub

