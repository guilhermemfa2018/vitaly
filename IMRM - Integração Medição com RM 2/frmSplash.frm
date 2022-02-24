VERSION 5.00
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{879115B9-8D7C-43CA-ADFE-8B489017BF42}#1.0#0"; "activelock1884.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login - IMRM"
   ClientHeight    =   6720
   ClientLeft      =   2325
   ClientTop       =   2850
   ClientWidth     =   11640
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Criar Banco e Tabelas no SQL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3615
      Left            =   5880
      TabIndex        =   11
      Top             =   2280
      Width           =   5655
      Begin IMRM.chameleonButton chameleonButton4 
         Height          =   495
         Left            =   3840
         TabIndex        =   25
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Salvar config"
         ENAB            =   0   'False
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplash.frx":15251
         PICN            =   "frmSplash.frx":1526D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton chameleonButton2 
         Height          =   495
         Left            =   2040
         TabIndex        =   24
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Criar tabelas"
         ENAB            =   0   'False
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplash.frx":15F47
         PICN            =   "frmSplash.frx":15F63
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton chameleonButton1 
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Criar banco"
         ENAB            =   0   'False
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplash.frx":16C3D
         PICN            =   "frmSplash.frx":16C59
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informações do DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   5295
         Begin IMRM.chameleonButton chameleonButton3 
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Verifica conexão"
            ENAB            =   0   'False
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSplash.frx":17933
            PICN            =   "frmSplash.frx":1794F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2520
            TabIndex        =   5
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Servidor:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Banco:"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Senha:"
            Height          =   255
            Left            =   2520
            TabIndex        =   14
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl+Shift+F12 configura Servidor - SQL Server"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   5415
      End
   End
   Begin IMRM.chameleonButton cmdRegistrarAgora 
      Height          =   495
      Left            =   9240
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Registre-se"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSplash.frx":18629
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Label7 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   7320
      TabIndex        =   19
      Top             =   3000
      Width           =   4215
      Begin IMRM.chameleonButton cmdSenha 
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "frmSplash.frx":18645
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin IMRM.chameleonButton cmdSenha 
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Login"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplash.frx":18661
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtCadastro 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Nome do usuário"
         ToolTipText     =   "Nome do usuário"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtCadastro 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   5760
   End
   Begin activelock1884.ActiveLock aLock 
      Left            =   3480
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   820
      SoftwareName    =   "SGCH"
      SoftwarePassword=   "2001"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
      LockToHardDrive =   0   'False
      LockToWindowsSerial=   -1  'True
      LockToRandomNumber=   -1  'True
      LockToComputerName=   0   'False
      LockToMACAddress=   0   'False
      UseDataLock     =   0   'False
      RegistryPath    =   "ActiveLock"
      RegistryKey     =   "VB and VBA Program Settings"
      RegistryName    =   "MyRegName"
      RegistryHive    =   "HKLM"
      LockToCustomString=   ""
      HashAlgorithm   =   0
      RegCounterKey   =   "Counter"
      RegLiberationKey=   "LiberationKey"
      RegLastRunDateKey=   "LastRunDate"
      RegInitialRunDateKey=   "InitialRunDate"
      RegRandomKey    =   "RandomKey"
      EncKey          =   "Default"
      RegEncKey       =   -1  'True
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6720
      Top             =   6120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6360
      Top             =   6120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6000
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5280
      Top             =   6120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5640
      Top             =   6120
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde! Carregando configurações do DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   6360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   5010
      Width           =   4935
   End
   Begin VB.Label lbldiasquefaltampararegistrar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "dias para a aplicação EXPIRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   4935
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage5 
      Height          =   660
      Left            =   240
      Top             =   4920
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1164
      Image           =   "frmSplash.frx":1867D
      Props           =   5
   End
   Begin AlphaImageControl.aicAlphaImage imgDemo 
      Height          =   450
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5280
      _ExtentX        =   9287
      _ExtentY        =   794
      Image           =   "frmSplash.frx":18E79
      Props           =   5
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   1800
      Left            =   7200
      Top             =   360
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   3175
      Image           =   "frmSplash.frx":1A4CF
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage3 
      Height          =   5640
      Left            =   120
      Top             =   -480
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   9948
      Image           =   "frmSplash.frx":2734D
      Props           =   5
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Const WM_SYSCOMMAND = &H112

Dim contsenha As Integer

Private Sub chameleonButton1_Click()
    PoeTempoMouse
    CriarBancoDeDadosADO
    TiraTempoMouse
End Sub

Private Sub chameleonButton2_Click()
    PoeTempoMouse
    CriarTabelasADO
    TiraTempoMouse
End Sub

Private Sub chameleonButton3_Click()
    PoeTempoMouse
    sServerName = Combo1.Text
    sDatabaseName = Combo2.Text
    sUsuName = Text1.Text
    sSenhaDB = Text2.Text
    sSGBD = 1
    'If Option1.Value = True Then sSGBD = 1
    'If Option2.Value = True Then sSGBD = 2
    Conexao
    TiraTempoMouse
End Sub

Private Sub chameleonButton4_Click()
On Error Resume Next
    PoeTempoMouse
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    
    sSGBD = 1
    'If Option1.Value = True Then sSGBD = 1
    'If Option2.Value = True Then sSGBD = 2
    
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sServerName", Combo1.Text 'Chave com o nome do Servidor
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sDatabaseName", Combo2.Text 'Chave com o nome do Banco
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sUsuName", Text1.Text '
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sSenhaDB", Text2.Text '
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sSGBD", sSGBD '
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sOFFLINE", "N" '
    
    'Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sLogoEmpresa", Combo2.Text 'Chave com o nome do Banco
    DesConfServer
    Label5.Caption = "Ctrl+Shift+F12 configura dados Servidor"
    TiraTempoMouse
    Msgbox "dados gravados com sucesso", vbInformation, "IMRM"
    
    CarregaDadosEmpresa
    aicAlphaImage1.Tag = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sLogoEmpresa")
    aicAlphaImage1.LoadImage_FromFile (aicAlphaImage1.Tag)
End Sub

Private Sub cmdRegistrarAgora_Click()
    frmRegistro.Show 1
    If varGlobal = "reiniciar" Then
        Msgbox "A aplicação será fechada. Inicie-a novamente"
        End
    End If
End Sub

Private Sub Combo2_DropDown()
On Error GoTo Err
    sServerName = Combo1.Text
    sDatabaseName = "master"
    Conexao
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [master].[dbo].[sysdatabases] ", cnBanco, 3, 3
    Combo2.Clear
    Do Until rs.EOF
        Combo2.AddItem rs("name")
        rs.MoveNext
    Loop
    Exit Sub
Err:
    Msgbox "O DB que esta tentando acessar não é gerenciado por essa aplicação"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Função q usa combinação de teclas para chamar outra função
    Dim TeclaSft        As Boolean
    Dim TeclaCtr        As Boolean
    TeclaSft = (Shift And vbShiftMask) > 0
    TeclaCtr = (Shift And vbCtrlMask) > 0
    If TeclaSft And TeclaCtr And KeyCode = vbKeyF12 Then
        If Frame1.Height = 15 Then
            Label8.Visible = True
            Timer5.Enabled = True
        Else
            DesConfServer
        End If
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Timer6_Timer()
'    atualizaVersao
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmdsenha_Click (1)
'    If KeyCode = &H71 Then
'        If Frame1.Height = 15 Then HabConfServer Else DesConfServer
'    End If
End Sub

Private Sub Cmdsenha_Click(Index As Integer)
'On Error GoTo TrataErro
    If txtCadastro(0).Enabled = True Then txtCadastro(0).SetFocus
    Select Case Index
        Case 0
            If Msgbox("Deseja encerrar a aplicação", vbQuestion + vbYesNo, "IMRM") = vbYes Then
                On Error Resume Next
                Kill "*.tmp"
                End
            End If
            If txtCadastro(0).Enabled = True Then txtCadastro(0).SetFocus
        Case 1
            If Conexao = False Then
                Exit Sub
            End If
'            Conexao
            
            Dim rsSenha As ADODB.Recordset
            Dim sql As String
            Set rsSenha = New ADODB.Recordset
            
            sql = "select a.codcoligada from tbsenha as a Where a.usuario= '" & txtCadastro(0).Text & "' and a.senha=" & " '" & txtCadastro(1).Text & "'"
            rsSenha.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
            If Not rsSenha.EOF Then vCodcoligada = rsSenha.Fields(0)
            rsSenha.Close
            
            sql = "Select a.usuario,a.senha,a.codigo,a.codcoligada,b.codigo,b.nome,b.email,b.codgrupo,b.altlogin,b.ativo,b.codcoligada,c.codigo,c.descricao,c.ativo,c.codcoligada,b.codcoligadatotvs,b.codven,b.nomeven,b.codusuariototvs,d.STATUS from tbSenha  as a inner join tbusuarios as b on a.codcoligada = '" & vCodcoligada & "' and a.codigo = b.codigo inner join tbgrupo as c on b.codgrupo = c.codigo left join CORPORE_ADM.dbo.GUSUARIO as d on b.codven COLLATE SQL_Latin1_General_CP1_CI_AS = d.CODUSUARIO Where a.usuario=" & " '" & txtCadastro(0).Text & "' and a.senha=" & " '" & txtCadastro(1).Text & "'"
            'sql = "Select * from tbSenha, tbusuarios, tbconfgrupo Where tbsenha.usuario=" & " '" & txtCadastro(0).Text & "'" & _
            '"and tbsenha.senha=" & " '" & txtCadastro(1).Text & "'" & "and tbsenha.codigo = tbUsuarios.codigo and tbconfgrupo.idgrupo = tbUsuarios.codgrupo"
            rsSenha.Open sql, cnBanco, adOpenKeyset, adLockReadOnly
            If rsSenha.RecordCount > 0 Then
                If rsSenha.Fields(19) = 0 Then
                    Msgbox "Usuário encontra-se desativado no RM"
                    Exit Sub
                End If
                CodUsu = rsSenha.Fields(2)
                NomUsu = rsSenha.Fields(5)
                GrupoUsu = rsSenha.Fields(12)
                vLogin = rsSenha.Fields(0)
                If Not IsNull(rsSenha.Fields(15)) Then vCodColigadaRM = Val(Mid$(rsSenha.Fields(15), 1, 6)) 'Código da coligada RM
                If Not IsNull(rsSenha.Fields(16)) Then vCodVenRM = rsSenha.Fields(16) 'codven do RM
                If Not IsNull(rsSenha.Fields(17)) Then vNomeVenRM = rsSenha.Fields(17) 'nomeven do RM
                If Not IsNull(rsSenha.Fields(18)) Then vCodUsuarioRM = rsSenha.Fields(18) 'codusuario do RM
                
'                FimAprop = rsSenha.Fields(18)
                'If Not IsNull(rsSenha.Fields(19)) Then vCodColigada = rsSenha.Fields(19)
                If rsSenha.Fields(0) <> txtCadastro(0) Then
                    Msgbox "Nome de usuário inválido"
                    Exit Sub
                End If
                If rsSenha.Fields(1) <> txtCadastro(1) Then
                    Msgbox "Senha inválida"
                    Exit Sub
                End If
                If rsSenha.Fields(8) = 1 Then
                    Msgbox "Sua senha expirou. Você precisa especificar uma nova senha para efeturar logon", vbQuestion, "Logon"
                    frmSplash.Tag = txtCadastro(0).Text
                    frmAlteraSenha.Show 1
                    If rsSenha.Fields(16) = 1 Then
                        Exit Sub
                    End If
                End If
                
                XCodGrp = rsSenha.Fields(7)
                '-----------
                rsSenha.Close
                Set rsSenha = Nothing
                
                If Not aLock.RegisteredUser Then
                    frmSplash.Visible = False
                Else
                    Unload frmSplash
                End If
                
                chamaParametro
                CarregaDadosEmpresa
                'atualizaMP 'Atualiza STATUS das programações de Métodos e Processos
                Principal.Show
                On Error GoTo TrataErro1 'ChecaNiver
                    
                Timer6.Enabled = True
                If vAvisos = "S" Then frmAvisos.Show 1
            Else
                Msgbox "Nome ou Senha do usuario inválido", vbInformation, "IMRM"
                txtCadastro(1).Text = ""
                txtCadastro(0).Text = ""
                txtCadastro(0).SetFocus
                contsenha = contsenha + 1
                rsSenha.Close
                Set rsSenha = Nothing
                If contsenha > 2 Then
                    cnBanco.Close
                    Set cnBanco = Nothing
                    End
                End If
            End If
        End Select
        Exit Sub
TrataErro1:
    Resume Next
TrataErro:
    HabConfServer
    Msgbox "Quantidade de Menus incompatíveis, entre no formulários de GRUPOS e salves todos os grupos"
    Principal.Show
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo Err1
    apontaLV = 100
    If App.PrevInstance = True Then
        End
    End If
    posBoxConfDB
    Dim Reg As Object
    Label6.Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    Set Reg = CreateObject("wscript.shell")
    
    If Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sPathIMRM") = "" Then
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sPathIMRM", App.Path & "\IMRM.exe" 'Chave com o nome do Servidor
    End If
    
    If Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sServerName") <> "" Then
        vIntegraOffline = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sOFFLINE")
        If vIntegraOffline = "S" Then
            vCaminhoReg = "HKEY_LOCAL_MACHINE\Software\IMRM\OFFLINE\"
        Else
            vCaminhoReg = "HKEY_LOCAL_MACHINE\Software\IMRM\"
        End If
        Combo1.Text = Reg.regread(vCaminhoReg & "sServerName")
        Combo2.Text = Reg.regread(vCaminhoReg & "sDatabaseName")
        Text1.Text = Reg.regread(vCaminhoReg & "sUsuName")
        Text2.Text = Reg.regread(vCaminhoReg & "sSenhaDB")
        
        If Combo1.Text = "" Then
            Combo1.Text = Reg.regread(vCaminhoReg & "sServerName")
            Combo2.Text = Reg.regread(vCaminhoReg & "sDatabaseName")
            Text1.Text = Reg.regread(vCaminhoReg & "sUsuName")
            Text2.Text = Reg.regread(vCaminhoReg & "sSenhaDB")
        End If
        
        aicAlphaImage1.Tag = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sLogoEmpresa")
        
        sServerName = Combo1.Text
        sDatabaseName = Combo2.Text
        sUsuName = Text1.Text
        sSenhaDB = Text2.Text
        
        'If Option1.Value = True Then sSGBD = 1
        'If Option2.Value = True Then sSGBD = 2
        aicAlphaImage1.LoadImage_FromFile (aicAlphaImage1.Tag)
        Label5.Caption = "F2 configura dados Servidor"
    End If
    
        'chamaParametro
        'If vBancoSAP <> "" Then
        '    ConexaoSAP
        '    CompoeComboSQL Combo3, "select a.CODLOC + ' - '+ a.NOME from tloc as a where a.CODCOLIGADA = 1 and a.INATIVO = 0"
        'End If
    
    '>>>>>>>>>>>
    'Registro 'remover para Orthoflex
    Exit Sub
Err1:
    If Err.Number = -2147024894 Then
        vIntegraOffline = "N"
        Resume Next
        Exit Sub
    End If
    
    HabConfServer
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sPathIMRM", App.Path & "\IMRM.exe" 'Chave com o nome do Servidor
    Reg.RegWrite vCaminhoReg & "sServerName", "" 'Chave com o nome do Servidor
    Reg.RegWrite vCaminhoReg & "sDatabaseName", "" 'Chave com o nome do Banco
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sLogoEmpresa", "" 'Logo da empresa
    Reg.RegWrite vCaminhoReg & "sUsuName", "" 'Logo da empresa
    Reg.RegWrite vCaminhoReg & "sSenhaDB", "" 'Logo da empresa
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sSGBD", "" 'Logo da empresa
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sOFFLINE", "N" '
    Label5.Caption = "Especifique dados de CONEXÃO"
    Exit Sub
End Sub

Private Sub posBoxConfDB()
    Frame1.Width = 975
    Frame1.Height = 15
    Frame1.Top = 120
    Frame1.Left = 10080
End Sub

Private Sub CarregaDadosEmpresa()
On Error GoTo Err
    Dim rsEmpresa As New ADODB.Recordset
    Dim sqlEmpresa As String
    Dim rsConfEmail As New ADODB.Recordset
    Dim sqlConfEmail As String
    Dim vLogoEmp As String
    
    sqlEmpresa = "Select *,CONVERT (VARCHAR, CURRENT_TIMESTAMP,103) as dataServidor from tbDadosEmpresa where codcoligada = '" & vCodcoligada & "'"
    rsEmpresa.Open sqlEmpresa, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsEmpresa.EOF Then
        NomeEmpresa = rsEmpresa.Fields(0)
        vLogoEmp = rsEmpresa.Fields(12)
        vDataDoBanco = rsEmpresa.Fields(16)
        SerieEmpresa = rsEmpresa.Fields(15)
    End If
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sLogoEmpresa", vLogoEmp 'Logo da empresa
    Set Reg = Nothing
    sqlConfEmail = "Select * from tbConfEmail where codcoligada = '" & vCodcoligada & "'"
    rsConfEmail.Open sqlConfEmail, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsConfEmail.EOF Then
        vSMTP = rsConfEmail.Fields(0)
        vUsuEmail = rsConfEmail.Fields(1)
        vSenhaEmail = rsConfEmail.Fields(2)
        vPorta = rsConfEmail.Fields(4)
        If rsConfEmail.Fields(5) = 0 Then
            vSSL = False
        Else
            vSSL = True
        End If
        vSMTPAutentic = rsConfEmail.Fields(5)
    End If
    rsConfEmail.Close
    If rsEmpresa.RecordCount > 0 Then rsEmpresa.Update
    rsEmpresa.Close
    Set rsEmpresa = Nothing
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub CarregaCombo()
    On Error GoTo Err
    
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    
    Dim dmoServer As SQLDMO.SQLServer
    Dim dmoApp As SQLDMO.Application
    Dim dmoNameList As SQLDMO.NameList
    Dim i As Integer
    Set dmoServer = New SQLDMO.SQLServer
    Set dmoApp = dmoServer.Application
    Set dmoNameList = dmoApp.ListAvailableSQLServers()
    Combo1.Clear
    For i = 0 To dmoNameList.Count - 1
        If dmoNameList(i) = "(local)" Then
            Dim PCName As String
            Dim P As Long
            P = NameOfPC(PCName)
            Combo1.AddItem PCName
        Else
            Combo1.AddItem dmoNameList(i)
        End If
    Next i
    Combo1.Text = Reg.regread(vCaminhoReg & "sServerName")
    Combo2.Text = Reg.regread(vCaminhoReg & "sDatabaseName")
        
    Text1.Text = Reg.regread(vCaminhoReg & "sUsuName")
    Text2.Text = Reg.regread(vCaminhoReg & "sSenhaDB")
    
    If Combo1.Text = "" Then
        Combo1.Text = Reg.regread(vCaminhoReg & "sServerName")
        Combo2.Text = Reg.regread(vCaminhoReg & "sDatabaseName")
        Text1.Text = Reg.regread(vCaminhoReg & "sUsuName")
        Text2.Text = Reg.regread(vCaminhoReg & "sSenhaDB")
    End If
    
    'If Reg.RegRead("HKEY_LOCAL_MACHINE\Software\IMRM\sSGBD") = 1 Then
    '    Option1.Value = True
    'Else
    '    Option2.Value = True
    'End If
    Exit Sub
Err:
    If Err.Number = 429 Or Err.Number = 91 Then
        'Msgbox "Aqui-1"
        
        'Resume Next
        Exit Sub
    End If
    If Err.Number <> -2147024894 Then
        Combo1.Text = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sServerName")
        Combo2.Text = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sDatabaseName")
        
        Text1.Text = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sUsuName")
        Text2.Text = Reg.regread("HKEY_LOCAL_MACHINE\Software\IMRM\sSenhaDB")
        'Msgbox "Aqui-2"
        
    End If
    
    If Combo1.Text = "" Then
        If sServerName <> "" Then
            Reg.RegWrite "HKEY_CURRENT_USER\Software\IMRM\" & "sPathIMRM", App.Path & "\IMRM.exe" 'Chave com o nome do Servidor
            Reg.RegWrite vCaminhoReg & "sServerName", "" 'Chave com o nome do Servidor
            Reg.RegWrite vCaminhoReg & "sDatabaseName", "" 'Chave com o nome do Banco
            Reg.RegWrite "HKEY_CURRENT_USER\Software\IMRM\" & "sLogoEmpresa", "" 'Logo da empresa
            Reg.RegWrite vCaminhoReg & "sUsuName", "" 'Logo da empresa
            Reg.RegWrite vCaminhoReg & "sSenhaDB", "" 'Logo da empresa
            Reg.RegWrite "HKEY_CURRENT_USER\Software\IMRM\" & "sSGBD", "" 'Logo da empresa
            Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sOFFLINE", "N" '
        
            Combo1.Text = Reg.regread(vCaminhoReg & "sServerName")
            Combo2.Text = Reg.regread(vCaminhoReg & "sDatabaseName")
            Text1.Text = Reg.regread(vCaminhoReg & "sUsuName")
            Text2.Text = Reg.regread(vCaminhoReg & "sSenhaDB")
        End If
        'Msgbox "Aqui-3"
    End If
    Exit Sub
End Sub

Public Function CriarBancoDeDadosADO() As Boolean
On Error GoTo Err1
    sServerName = Combo1.Text
    sDatabaseName = Combo2.Text
    sUsuName = Text1.Text
    sSenhaDB = Text2.Text
    sSGBD = 1
    'If Option1.Value = True Then sSGBD = 1
    'If Option2.Value = True Then sSGBD = 2
    
    Set oConn = New ADODB.Connection
    
    ' Cria Banco
'    oConn.Open "Provider=SQLOLEDB;Data Source=" & sServerName & ";User ID=sa;Password=;"
    'If Option1.Value = True Then
        oConn.Open "Provider=SQLOLEDB;Data Source=" & sServerName & ";User ID=" & sUsuName & ";Password=" & sSenhaDB & ";"
    'Else
    'End If
    oConn.Execute "CREATE DATABASE " & sDatabaseName
    
    oConn.Close
    Set oConn = Nothing
    
    Msgbox "Banco criado com sucesso", vbInformation, "IMRM"
    Exit Function
Err1:
    Msgbox "(ADO) Erro ao criar banco de dados: " & vbCrLf & Err.Number & " - DB já Existe - " & Err.Description, 16, "Mensagem de erro"
    Exit Function
End Function

Private Sub HabConfServer()
    If Frame1.Height = 15 Then Expande Else Recolhe
    Frame1.Enabled = True
    chameleonButton1.Enabled = True
    chameleonButton2.Enabled = True
    chameleonButton3.Enabled = True
    chameleonButton4.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = True
    Label5.Caption = "Especifique dados de CONEXÃO"
End Sub

Private Sub DesConfServer()
    If Frame1.Height = 15 Then Expande Else Recolhe
    Frame1.Enabled = True
    chameleonButton1.Enabled = False
    chameleonButton2.Enabled = False
    chameleonButton3.Enabled = False
    chameleonButton4.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    'Label8.Visible = False
    Label5.Caption = "Especifique dados de CONEXÃO"
End Sub

Private Sub PoeTempoMouse()
    frmSplash.MousePointer = 11
    chameleonButton1.MousePointer = 11
    chameleonButton2.MousePointer = 11
    chameleonButton3.MousePointer = 11
    chameleonButton4.MousePointer = 11
    cmdSenha(0).MousePointer = 11
    cmdSenha(1).MousePointer = 11
End Sub
Private Sub TiraTempoMouse()
    frmSplash.MousePointer = 0
    chameleonButton1.MousePointer = 0
    chameleonButton2.MousePointer = 0
    chameleonButton3.MousePointer = 0
    chameleonButton4.MousePointer = 0
    cmdSenha(0).MousePointer = 0
    cmdSenha(1).MousePointer = 0
End Sub

Private Sub Recolhe()
    Timer2.Enabled = True
    Timer4.Enabled = True
End Sub

Private Sub Expande()
    Timer3.Enabled = True
    Timer1.Enabled = True
End Sub

Private Sub MoveLabel7()

End Sub

Private Sub Timer1_Timer()
    If Frame1.Height <= 3615 Then
        Frame1.Width = 5655
        Frame1.Left = 5880
        Frame1.Height = Frame1.Height + 200
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
    If Frame1.Height > 15 Then
        Frame1.Height = Frame1.Height - 200
        Frame1.Width = 975
        Frame1.Left = 5880
    Else
        Timer2.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
    If Label7.Top > 1680 Then
        Label7.Top = Label7.Top - 200
        Label1.Top = Label1.Top - 200
        Label2.Top = Label2.Top - 200
        txtCadastro(0).Top = txtCadastro(0).Top - 200
        txtCadastro(1).Top = txtCadastro(1).Top - 200
        cmdSenha(1).Top = cmdSenha(1).Top - 200
        cmdSenha(0).Top = cmdSenha(0).Top - 200
    Else
        Timer3.Enabled = False
    End If
End Sub

Private Sub Timer4_Timer()
    If Label7.Top < 3000 Then
        Label7.Top = Label7.Top + 200
        Label1.Top = Label1.Top + 200
        Label2.Top = Label2.Top + 200
        txtCadastro(0).Top = txtCadastro(0).Top + 200
        txtCadastro(1).Top = txtCadastro(1).Top + 200
        cmdSenha(1).Top = cmdSenha(1).Top + 200
        cmdSenha(0).Top = cmdSenha(0).Top + 200
    Else
        Timer4.Enabled = False
    End If
End Sub

Private Sub Timer5_Timer()
    CarregaCombo
    HabConfServer
    Timer5.Enabled = False
    Label8.Visible = False
End Sub

Private Sub chamaParametro()
    Dim rsParametros As New ADODB.Recordset
    Dim sqlParametros As String
    Dim rsIntegra As New ADODB.Recordset
    Dim sqlIntegra As String
    Dim rsIntegraOff As New ADODB.Recordset
    Dim sqlIntegraOff As String
    
    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsParametros.EOF Then
        MediaGlobal = rsParametros.Fields(0)
        GeraIntr = rsParametros.Fields(1)
        GeraObri = rsParametros.Fields(4)
        If Not IsNull(rsParametros.Fields(3)) Then GeraLog = rsParametros.Fields(3)
        vAprovadoRest = rsParametros.Fields(2)
        If Not IsNull(rsParametros.Fields(5)) Then vIntegra = rsParametros.Fields(5)
        
        If Not IsNull(rsParametros.Fields(7)) Then vAvisos = rsParametros.Fields(7)
        
        If Not IsNull(rsParametros.Fields(8)) Then vCaminhoAtu = rsParametros.Fields(9)
        If Not IsNull(rsParametros.Fields(10)) Then vCalcExp = rsParametros.Fields(10)
        If Not IsNull(rsParametros.Fields(11)) Then vAfastDias = rsParametros.Fields(11)
        If Not IsNull(rsParametros.Fields(12)) Then vAfastTreiInt = rsParametros.Fields(12)
        If Not IsNull(rsParametros.Fields(13)) Then vAfastTreiObr = rsParametros.Fields(13)
        If Not IsNull(rsParametros.Fields(14)) Then vInicioAvOC = rsParametros.Fields(14)
        If Not IsNull(rsParametros.Fields(2)) Then vPeriodo = rsParametros.Fields(2)
        
        If vIntegra = "S" Then
            sqlIntegra = "Select * from tbintegracao where codcoligada = '" & vCodcoligada & "'"
            rsIntegra.Open sqlIntegra, cnBanco, adOpenKeyset, adLockReadOnly
            If Not rsIntegra.EOF Then
                vServerSAP = rsIntegra.Fields(3)
                vBancoSAP = rsIntegra.Fields(4)
                vUsuBancoTovs = rsIntegra.Fields(5)
                vSenhaBancoSAP = rsIntegra.Fields(6)
                'se estiver usando o banco FERRAMENTARIA_OFF vai usar as configuracoes abaixo
                If Combo2.Text = "FERRAMENTARIA_OFF" Then
                    vServerSAP = Combo1.Text ' = Reg.regread(vCaminhoReg & "sServerName")
                    vBancoSAP = "CORPORERM_OFF" '= Reg.regread(vCaminhoReg & "sDatabaseName")
                    vUsuBancoTovs = Text1.Text ' = Reg.regread(vCaminhoReg & "sUsuName")
                    vSenhaBancoSAP = Text2.Text ' = Reg.regread(vCaminhoReg & "sSenhaDB")
                End If
            End If
            rsIntegra.Close
            Set rsIntegra = Nothing
        End If
        
        If vIntegraOffline = "S" Then
            Dim Reg As Object
            Set Reg = CreateObject("wscript.shell")
            vServerSAP = Reg.regread(vCaminhoReg & "sServerName")
            vBancoSAP = Reg.regread(vCaminhoReg & "sDatabaseNameCORPORERM_OFF")
            vUsuBancoTovs = Reg.regread(vCaminhoReg & "sUsuName")
            vSenhaBancoSAP = Reg.regread(vCaminhoReg & "sSenhaDB")
            Set Reg = Nothing
        End If
        
        'sqlIntegraOff = "Select * from tbintegracaooffline"
        'rsIntegraOff.Open sqlIntegraOff, cnBanco, adOpenKeyset, adLockReadOnly
        'If Not rsIntegraOff.EOF Then
        '    If rsIntegraOff.Fields(0) = "S" Then
        '        vIntegraOffline = "S"
        '        vServerOffline = rsIntegraOff.Fields(1)
        '        vBancoOffline = rsIntegraOff.Fields(2)
        '        vUsuBancoOffline = rsIntegraOff.Fields(3)
        '        vSenhaBancoOffline = rsIntegraOff.Fields(4)
        '    End If
        'End If
        'rsIntegraOff.Close
        'Set rsIntegraOff = Nothing
        
    End If
    rsParametros.Close
    Set rsParametros = Nothing
    
    Dim rsVerLinguagem As New ADODB.Recordset
    Dim SqlVerLinguagem As String
    SqlVerLinguagem = "SELECT @@language"
    rsVerLinguagem.Open SqlVerLinguagem, cnBanco, adOpenKeyset, adLockReadOnly
    If rsVerLinguagem.Fields(0) = "us_english" Then vFormatoDatetime = "yyyy-mm-dd" Else vFormatoDatetime = "dd-mm-yyyy"
    rsVerLinguagem.Close
    Set rsVerLinguagem = Nothing
End Sub

Private Sub Registro()
    
    aLock.SoftwareName = "IMRM" '& Format(Now, "yy")
    aLock.SoftwarePassword = "2001" '& Format(Now, "yy")
'    aLock.SoftwareName = "IMRM" & Format(Now, "mm/yy")
'    aLock.SoftwarePassword = "2001" & Format(Now, "mm/yy")
    
    If Not aLock.RegisteredUser Then
        If aLock.LastRunDate > Now Then
            If Msgbox("Ouve alteração na data do Sistema, inferior a data que o mesmo foi registrado " _
            & vbCrLf & "O Programa deve ser reativado na data atual ou mude a data para a data superior " _
            & vbCrLf & "que o mesmo foi registrado.", vbOKOnly + vbInformation, "Data Alterada") = vbOK Then
                End
            End If
        End If
        
        imgDemo.Visible = True
        lbldiasquefaltampararegistrar.Visible = True
    Else
        imgDemo.Visible = False
        lbldiasquefaltampararegistrar.Visible = False
    End If
    
    Dim diasQueFaltaParaRegistrar As Integer
    
    diasQueFaltaParaRegistrar = 0
    diasQueFaltaParaRegistrar = 30 - (aLock.UsedDays)
    If diasQueFaltaParaRegistrar < 0 Then diasQueFaltaParaRegistrar = 0
    lbldiasquefaltampararegistrar = Str(diasQueFaltaParaRegistrar) & " " & lbldiasquefaltampararegistrar

    If Not aLock.RegisteredUser Then
        If diasQueFaltaParaRegistrar <= 0 Then
            lbldiasquefaltampararegistrar.Visible = True
            lbldiasquefaltampararegistrar = "Sua aplicação EXPIROU!"
            cmdSenha(1).Enabled = False
            txtCadastro(0).Enabled = False
            txtCadastro(1).Enabled = False
            cmdRegistrarAgora.Visible = True
        End If
    End If
End Sub

Private Sub atualizaVersao()
On Error GoTo Err
    Dim camIMRMso As String
    Dim shell1, strOS, strVerKey, strVersion
    Set shell1 = CreateObject("WScript.Shell")
    strOS = shell1.ExpandEnvironmentStrings("%OS%")
    If strOS = "Windows_NT" Then
        strVerKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
        strVersion = shell1.regread(strVerKey & "ProductName")
    Else
        strVerKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\"
        strVersion = shell1.regread(strVerKey & "ProductName")
    End If
    Set shell1 = Nothing
    
    If vCaminhoAtu = "" Then Exit Sub
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim caminhoIMRMAtu As String
    
    caminhoIMRMAtu = Mid$(vCaminhoAtu, 1, Len(vCaminhoAtu) - 15) & "IMRM.exe"
    
    'Verificar se o Executavel AtualizaIMRM.exe existe
    If Dir$(vCaminhoAtu) <> "" Then
        'Verificar se o Executavel AtualizaIMRM.exe existe
        If Dir$(caminhoIMRMAtu) <> "" Then
            camIMRMso = App.Path & "\IMRM.exe"
            If fso.GetFileVersion(caminhoIMRMAtu) > fso.GetFileVersion(camIMRMso) Then
                If Msgbox("Uma nova versão do sistema foi disponibilizada no REPOSITÓRIO. Deseja atualizar?", vbQuestion + vbYesNo, "IMRM") = vbYes Then
                    Shell vCaminhoAtu, vbNormalFocus
                End If
            End If
        End If
    End If
    Timer6.Enabled = False
    Exit Sub
Err:
    If Err.Number = 76 Then Msgbox "Executável de atualização não encontrado no caminho informado", vbCritical, "IMRM"
    Exit Sub
End Sub

'Substitui aspas simples por aspas duplas
Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 34
    End If
End Sub




