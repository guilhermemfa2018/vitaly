VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmGeraRelatório 
   Caption         =   "Imprimir Ordem de Serviço"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGeraRelatório.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Informe o diretório do DB "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
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
         Index           =   5
         Left            =   5760
         TabIndex        =   6
         ToolTipText     =   "Localizar arquivo"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtIntegra 
         Height          =   330
         Index           =   10
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Caminho onde será gravado o arquivo com os dados capturados no relógio de ponto"
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nº da Ordem de Serviço "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
      Begin VB.TextBox txtIntegra 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog cdlTXT3 
      Left            =   5880
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   6615
   End
End
Attribute VB_Name = "frmGeraRelatório"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCad_Click(Index As Integer)
On Error GoTo Err
        With cdlTXT3
            .DialogTitle = "Selecione um diretório"
            .InitDir = App.Path
            .FileName = "Selecione um diretório"
            .Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
            .Filter = "(Diretórios|*.~#~"
            .CancelError = True
            .ShowSave
            Caminho4 = .FileName
        End With
        txtIntegra(10) = CurDir + "\"
        sDatabaseName = txtIntegra(10)
        'sCompanyName = txtIntegra(11)
        gravaCaminhoNoRegedit
        carregaCaminhoDoRegedit
    conexaoDBF
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub Command1_Click()
    conexaoDBF
End Sub

Private Sub listarDadosDBF()
    Dim rslistarDadosDBF As New ADODB.Recordset
    Dim SqllistarDadosDBF As String
    
    SqllistarDadosDBF = SqllistarDadosDBF & "SELECT " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.NUMERO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.CLIENTE, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.CONTATO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.TELEFONE, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.EQUIPA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.MODELO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.MARCA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.ACESSORIOS, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SERIE, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SITUACAO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.DEFEITO1, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.DEFEITO2, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SERVICO1, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SERVICO2, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SERVICO3, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SERVICO4, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.SERVICO5, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.OBSERVA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.TECNICO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.HORAENT, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.DATAENT, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.HORASAI, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.DATASAI, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.MAO_OBRA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.DESCONTO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.TOTAL, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.VLRPROD, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " A.PECAS, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.EMPRESA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.RUA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.EMAIL, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.NUM, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.CID, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.UF, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.CEP, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.BAI, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " B.FAX AS TEL, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " C.E_LOGO AS LOGO, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " D.RUA AS CLI_RUA, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " D.BAI AS CLI_BAI, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " D.CID AS CLI_CID, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " D.UF AS CLI_UF, " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " D.CEP AS CLI_CEP " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & " " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & "FROM ORDEM AS A, DADOS AS B, LOGO AS C, CLIENTES AS D " & vbCrLf
    SqllistarDadosDBF = SqllistarDadosDBF & "WHERE"
    SqllistarDadosDBF = SqllistarDadosDBF & " A.NUMERO = " & vOS & " AND A.CODCLI = D.COD"
    
    rslistarDadosDBF.Open SqllistarDadosDBF, cnBancoDBF, adOpenKeyset, adLockReadOnly
    If rslistarDadosDBF.RecordCount > 0 Then
        MsgBox rslistarDadosDBF.RecordCount
    Else
        MsgBox "Nada Encontrado"
    End If
End Sub

Private Sub Command2_Click()
    vOS = txtIntegra(0)
    listarDadosDBF
End Sub

Private Sub Command3_Click()
    If txtIntegra(0).Text <> "" Then
        vOS = txtIntegra(0).Text
        'sCompanyName = txtIntegra(11)
        gravaCaminhoNoRegedit
        FCRCRelatorio.Show 1
    Else
        MsgBox "Informe um número de OS"
    End If
End Sub

Private Sub Form_Load()
    carregaCaminhoDoRegedit
    conexaoDBF
End Sub
