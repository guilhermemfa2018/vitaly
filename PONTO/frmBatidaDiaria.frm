VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatidaDiaria 
   Caption         =   "Importar Batidas"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4620
   Icon            =   "frmBatidaDiaria.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   11000
      Left            =   2520
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   14000
      Left            =   1080
      Top             =   4560
   End
   Begin VB.Frame Frame2 
      Caption         =   "Importar Batidas - Período"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   4095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   17000
         Left            =   2880
         Top             =   960
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "Importar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   141295617
         CurrentDate     =   41499
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   141361153
         CurrentDate     =   41499
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tratar Batidas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   4335
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "Localizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   11
         Top             =   840
         Width           =   3855
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "Ler"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   2160
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog cdlTXT 
         Left            =   3600
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text7 
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
         TabIndex        =   4
         Text            =   "SRV1002\CORPORERM"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text6 
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
         Left            =   2400
         TabIndex        =   3
         Text            =   "ZEUS"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text5 
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
         Text            =   "sa"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text4 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "vigamax"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   4335
   End
End
Attribute VB_Name = "frmBatidaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Caminho2 As String
Public vValidaImportacao As Boolean
Public vContaBatida As Integer

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        importaBatidas RemoveMask(DTPicker1.Value), RemoveMask(DTPicker2.Value)
    Case 11
        With cdlTXT
            .Filter = "(Arquivo *.TXT)|*.txt"
            .ShowOpen
            Caminho2 = .FileName
        End With
        Text1 = Caminho2
        If Text1.Text <> "" Then cmdCadastro(14).Enabled = True
    Case 14
        trataBatidas
    End Select
End Sub

Private Sub deletaDadosDoDia()
    Dim rsDeletaDados As New ADODB.Recordset
    Dim SqlDeletaDados As String
    
    SqlDeletaDados = "DELETE FROM TBPONTO WHERE DATABATIDA  = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & "'"
    rsDeletaDados.Open SqlDeletaDados, cnBanco, adOpenKeyset, adLockReadOnly
End Sub


Private Sub trataBatidas()
    Dim vID As String, vTipo As String, vData As String, vHora As String, vPis As String
    Dim X As Integer
    Dim F As Long
    Dim Linhas As Variant
    Dim i As Long
    Dim Tmp As String
    
    deletaDadosDoDia
    
    F = FreeFile
    vContaBatida = 0
    Open Text1.Text For Input As #F

    Tmp = Input(LOF(F), F)
    Close #F

    Linhas = Split(Tmp, Chr(10))
    For i = 0 To UBound(Linhas) - 1
        vTipo = Mid$(Linhas(i), 10, 1)
      
        If vTipo = 3 Then
            vID = Mid$(Linhas(i), 1, 9)
            vData = Mid$(Linhas(i), 11, 8)
            vHora = Mid$(Linhas(i), 19, 4)
            vPis = Mid$(Linhas(i), 23, 12)
            vData = formatData(vData)
            vHora = formatHora(vHora)
            
            insertDados vID, vData, vHora, vPis
        End If
    Next
    If excluiArquivos(RemoveMask(DTPicker1.Value)) = False Then
        MsgBox "Erro ao tentar excluir os arquivos txt na pasta c:\zeus\ponto", vbCritical, "Atenção"
    End If
    
    If conexaoFB = True Then
        Label1.Caption = "Importando horário do FlexJr"
        buscaHorariosFlexJunior
    Else
        'MsgBox "Erro ao tentar realizar conexão com o banco FLEXJUNIOR", vbCritical, "Atenção"
        Label1.Caption = "Removendo Ponto.exe da memoria"
        Timer3.Enabled = True
    End If
    
    'MsgBox "Dados importados com sucesso!", vbInformation, "SGCH"
End Sub

Private Function ValidaDados()
    ValidaDados = False
    Dim Y As Integer
    For Y = 0 To 3
        If colheDados(Y) = "" Then
            MsgBox "Erro de consistência na fonte de dados", vbCritical, "Ponto"
            Exit Function
        End If
    Next
    ValidaDados = True
End Function

Private Sub insertDados(vIdentificador As String, vDataBatida As String, vHoraBatida As String, vPISPASEP As String)
On Error Resume Next
    Dim rsEncontraChapaPeloPIS As New ADODB.Recordset
    Dim SqlEncontraChapaPeloPIS As String, vChapa As String
    
    Dim rsGravaDadosTratados As New ADODB.Recordset
    Dim SqlGravaDadosTratados As String
    
    SqlEncontraChapaPeloPIS = "SELECT A.CHAPA,A.PISPASEP FROM CORPORERM.DBO.PFUNC AS A WHERE A.CODCOLIGADA = 6 AND A.CODSITUACAO <> 'D' AND A.PISPASEP = '" & Val(vPISPASEP) & "'"
    rsEncontraChapaPeloPIS.Open SqlEncontraChapaPeloPIS, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsEncontraChapaPeloPIS.RecordCount > 0 Then
        vChapa = rsEncontraChapaPeloPIS.Fields(0)

        SqlGravaDadosTratados = "SELECT * FROM TBPONTO WHERE DATABATIDA  = '" & vDataBatida & "' AND CHAPA  = '" & vChapa & "'"
        rsGravaDadosTratados.Open SqlGravaDadosTratados, cnBanco, adOpenKeyset, adLockReadOnly
        If rsGravaDadosTratados.RecordCount = 0 Then
            rsGravaDadosTratados.Close
            vContaBatida = 1
            SqlGravaDadosTratados = "Insert into TBPONTO(ID,DATABATIDA,CHAPA,BATIDA1,CONTBATIDA) Values('" & vIdentificador & "','" & vDataBatida & "','" & vChapa & "','" & vHoraBatida & "', " & vContaBatida & ")"
            rsGravaDadosTratados.Open SqlGravaDadosTratados, cnBanco
        Else
            vContaBatida = rsGravaDadosTratados.Fields(9)
            rsGravaDadosTratados.Close
            
            
            'If vChapa = "00048" Then
            '    vChapa = "00048"
            'End If
            
            
            If vContaBatida = 1 Then
                SqlGravaDadosTratados = "Update TBPONTO SET BATIDA2 = '" & vHoraBatida & "', CONTBATIDA = " & vContaBatida + 1 & " WHERE DATABATIDA  = '" & vDataBatida & "' AND CHAPA  = '" & vChapa & "'"
            ElseIf vContaBatida = 2 Then
                SqlGravaDadosTratados = "Update TBPONTO SET BATDA3 = '" & vHoraBatida & "', CONTBATIDA = " & vContaBatida + 1 & " WHERE DATABATIDA  = '" & vDataBatida & "' AND CHAPA  = '" & vChapa & "'"
            ElseIf vContaBatida = 3 Then
                SqlGravaDadosTratados = "Update TBPONTO SET BATIDA4 = '" & vHoraBatida & "', CONTBATIDA = " & vContaBatida + 1 & " WHERE DATABATIDA  = '" & vDataBatida & "' AND CHAPA  = '" & vChapa & "'"
            ElseIf vContaBatida = 4 Then
                SqlGravaDadosTratados = "Update TBPONTO SET BATIDA5 = '" & vHoraBatida & "', CONTBATIDA = " & vContaBatida + 1 & " WHERE DATABATIDA  = '" & vDataBatida & "' AND CHAPA  = '" & vChapa & "'"
            ElseIf vContaBatida = 5 Then
                SqlGravaDadosTratados = "Update TBPONTO SET BATIDA6 = '" & vHoraBatida & "', CONTBATIDA = " & vContaBatida + 1 & " WHERE DATABATIDA  = '" & vDataBatida & "' AND CHAPA  = '" & vChapa & "'"
            End If
            rsGravaDadosTratados.Open SqlGravaDadosTratados, cnBanco
        End If
    End If
    rsEncontraChapaPeloPIS.Close
    Set rsEncontraChapaPeloPIS = Nothing
    Exit Sub
End Sub

Private Function excluiArquivos(vNomeArquivo As String)
On Error GoTo Err
    excluiArquivos = True
    Dim vComando As String
    Kill "C:\ZEUS\PONTO\" & vNomeArquivo & ".TXT"
    'End
    Exit Function
Err:
    excluiArquivos = False
    Exit Function
End Function

Private Function importaBatidas(vDataIni As String, vDataFim As String)
On Error GoTo Err
    'Captura batidas do relogio de ponto da Vitaly (Localização: Refeitorio)
    vValidaImportacao = True
    'Dim vComando As String
    
    'CAPTURA DADOS DAS BATIDAS NO RELOGIO DE PONTO
    vComando = "ckrep /c """ & vCaminhoDadosCapturadoRelogio & """ """ & vDataIni & ".TXT"" /net /nr" & vIDRelogio & " /s" & vPassRelogio & " /cpf" & vCPFResponsavel & " /dti" & vDataIni & "." & vDataFim & " /ip" & vIPRelogio & ""
'    vComando = "ckrep /c ""C:\ZEUS\PONTO\"" """ & vDataIni & ".TXT"" /net /nr00008003930105288 /s10000001 /cpf61523372672 /dti" & vDataIni & "." & vDataFim & " /ip192.168.0.163"
    
    
    
    
    'CAPTURA DADOS DOS FUNCIONARIOS NO RELOGIO DE PONTO
    'vComando = "ckrep /l ""C:\ZEUS\PONTO\Coleta\teste.txt""  /net /nr00008003930105288 /s10000001 /cpf61523372672 /dti01102021.01102021 /ip192.168.0.163"
    
    Shell vComando
    Timer1.Enabled = True
    Exit Function
Err:
    vValidaImportacao = False
    Exit Function
End Function

Private Sub Timer1_Timer()
    If vValidaImportacao = True Then
        'MsgBox "Batidas capturadas com sucesso", vbInformation, "SUCESSO"
        
        Label1.Caption = "Tratando batidas"
        Text1.Text = vCaminhoDadosCapturadoRelogio & RemoveMask(DTPicker1.Value) & ".TXT"
        trataBatidas
        
        'Timer2.Enabled = True
    Else
        MsgBox "Ocorreu um erro durante a captura das batidas", vbCritical, "ATENÇÃO"
    End If
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    'Label1.Caption = "Tratando batidas"
    'Text1.Text = "C:\ZEUS\PONTO\" & RemoveMask(DTPicker1.Value) & ".TXT"
    'trataBatidas
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Dim Comando As String
    Comando = "TASKKILL -F -IM Ponto"
    Shell Comando
    End
End Sub

Private Sub buscaHorariosFlexJunior()
    Dim rsBuscaHorariosFlexJunior As New ADODB.Recordset
    Dim SqlBuscaHorariosFlexJunior As String
    
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "SELECT A.FUN_CODIGO " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     , A.FUN_DESCRICAO " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     , B.HOR_ENTRADA " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     , B.HOR_INICIO_REFEICAO " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     , B.HOR_FINAL_REFEICAO " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     , B.HOR_SAIDA " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "FROM CADFUN AS A " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "JOIN TABHOR AS B ON " & vbCrLf
    'SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     A.FUN_QUADRO_HORARIO = B.HOR_CODIGO"
    'rsBuscaHorariosFlexJunior.Open SqlBuscaHorariosFlexJunior, cnBancoFlexJunior, adOpenKeyset, adLockReadOnly

    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "SELECT " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.FUN_CODIGO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.DESCRICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        WHEN D.HRC_ENTRADA IS NOT NULL THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            D.HRC_ENTRADA " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            E.QDO_ENTRADA_SEGUNDA " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    END AS HOR_ENTRADA, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        WHEN E.QDO_INICIO_REF_SEGUNDA IS NOT NULL THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            E.QDO_INICIO_REF_SEGUNDA " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            '-' " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    end  HOR_INICIO_REFEICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        WHEN E.QDO_FINAL_REF_SEGUNDA IS NOT NULL THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            E.QDO_FINAL_REF_SEGUNDA " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            '-' " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    end  HOR_FINAL_REFEICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        WHEN D.HRC_SAIDA IS NOT NULL THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                WHEN  D.HRC_ENTRADA <> '23:00' THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    SUBSTRING(D.HRC_SAIDA FROM 1 FOR 5) " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    CAST(CAST(CAST(SUBSTRING(D.HRC_ENTRADA FROM 1 FOR 2) AS INT) + (CAST(24 - CAST(SUBSTRING(D.HRC_ENTRADA FROM 1 FOR 2) AS INT) AS INT)+CAST(SUBSTRING(D.HRC_SAIDA FROM 1 FOR 2) AS INT)) AS INT) AS VARCHAR(2))  || ':00' " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            END " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            CAST(E.QDO_SAIDA_SEGUNDA AS VARCHAR(5)) " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    END AS HOR_SAIDA " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "FROM( " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    SELECT " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        LPad(CAST(A.FUN_CODIGO AS INT), 5, 0) AS FUN_CODIGO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_DESCRICAO AS DESCRICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_QUADRO_HORARIO AS QUADRO_HORARIO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        B.QDO_DESCRICAO AS QDO_DESCRICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        MAX(C.HIS_DATA) AS DT_CICLO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            WHEN (A.FUN_QUADRO_HORARIO = 1) or (A.FUN_QUADRO_HORARIO IS NULL)  THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                '-' " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "            ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                CASE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    WHEN ((FLOOR((CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)/21 + 1) * 21) - 21 - (CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)) * - 1 + 1 IN(1,2,3,4,5,6,7) THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                        1 " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    WHEN ((FLOOR((CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)/21 + 1) * 21) - 21 - (CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)) * - 1 + 1 IN(8,9,10,11,12,13,14) THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                        2 " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    WHEN ((FLOOR((CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)/21 + 1) * 21) - 21 - (CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)) * - 1 + 1 IN(15,16,17,18,19,20,21) THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                        3 " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    WHEN ((FLOOR((CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)/21 + 1) * 21) - 21 - (CAST(current_timestamp - MAX(C.HIS_DATA)-1 AS INT)+1)) * - 1 + 1 IS NULL THEN " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                        '-' " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                    ELSE " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                        '-' " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "                END " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        END AS CICLO " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    FROM cadfun AS A " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    INNER JOIN QDOHORAR AS B ON " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_QUADRO_HORARIO = B.QDO_CODIGO " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    LEFT JOIN HISCON AS C ON " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_CODIGO  = C.HIS_FUNCIONARIO " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    GROUP BY " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_CODIGO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_DESCRICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        A.FUN_QUADRO_HORARIO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "        B.QDO_DESCRICAO " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & ") AS CICLO " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "LEFT JOIN HREVCICL AS D ON " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     CICLO.QUADRO_HORARIO = D.HRC_REVEZAMENTO AND " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "     'CICLO 0' || CICLO.CICLO = D.HRC_CICLO AND " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    D.HRC_ENTRADA IS NOT NULL " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "LEFT JOIN QDOHORAR AS E ON " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.QUADRO_HORARIO = E.QDO_CODIGO " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "GROUP BY " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.FUN_CODIGO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.DESCRICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.QUADRO_HORARIO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.QDO_DESCRICAO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.DT_CICLO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    CICLO.CICLO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    D.HRC_CICLO, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    D.HRC_ENTRADA, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    D.HRC_SAIDA, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    E.QDO_ENTRADA_SEGUNDA, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    E.QDO_INICIO_REF_SEGUNDA, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    E.QDO_FINAL_REF_SEGUNDA, " & vbCrLf
    SqlBuscaHorariosFlexJunior = SqlBuscaHorariosFlexJunior & "    E.QDO_SAIDA_SEGUNDA"
    rsBuscaHorariosFlexJunior.Open SqlBuscaHorariosFlexJunior, cnBancoFlexJunior, adOpenKeyset, adLockReadOnly

    If rsBuscaHorariosFlexJunior.RecordCount > 0 Then
        'MsgBox "FORAM ENCONTRADOS: " & rsBuscaHorariosFlexJunior.RecordCount & " REGISTROS"
        Dim vSaida As String
        Do While Not rsBuscaHorariosFlexJunior.EOF
            If Not IsNull(rsBuscaHorariosFlexJunior.Fields(5)) And rsBuscaHorariosFlexJunior.Fields(5) <> "" Then
                vSaida = acertaHoraSaida(rsBuscaHorariosFlexJunior.Fields(2), rsBuscaHorariosFlexJunior.Fields(5))
            Else
                vSaida = acertaHoraSaida(rsBuscaHorariosFlexJunior.Fields(2), rsBuscaHorariosFlexJunior.Fields(4))
            End If
            exportaHorarioParaZeus Format(rsBuscaHorariosFlexJunior.Fields(0), "00000"), rsBuscaHorariosFlexJunior.Fields(2), vSaida, 6
            rsBuscaHorariosFlexJunior.MoveNext
        Loop
    End If
    rsBuscaHorariosFlexJunior.Close
    cnBancoFlexJunior.Close
    Timer3.Enabled = True
End Sub

Private Function acertaHoraSaida(vHoraEntrada As String, vHoraSaida As String)
    Dim rsAcertaHoraSaida As New ADODB.Recordset
    Dim SqlAcertaHoraSaida As String
    Dim vDiferenca As String
    
    'SqlAcertaHoraSaida = SqlAcertaHoraSaida & "DECLARE @DIFERENCA AS VARCHAR(10) " & vbCrLf
    'SqlAcertaHoraSaida = SqlAcertaHoraSaida & "SELECT @DIFERENCA = REPLICATE('0', 2 - LEN(CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraSaida & "')) - CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraEntrada & "')))) + RTrim(CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraSaida & "')) - CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraEntrada & "'))) + ':00'" & vbCrLf
    'SqlAcertaHoraSaida = SqlAcertaHoraSaida & "SELECT " & vbCrLf
    'SqlAcertaHoraSaida = SqlAcertaHoraSaida & "   CONVERT(VARCHAR(10),FORMAT(CONVERT(DATETIME,FORMAT(getdate(), 'yyyy-MM-dd') + ' ' + '" & vHoraEntrada & "')  + @DIFERENCA,'yyyy-MM-dd'), 112) AS DATA_SAIDA, " & vbCrLf
    'SqlAcertaHoraSaida = SqlAcertaHoraSaida & "   CONVERT(VARCHAR(5),CONVERT(DATETIME,FORMAT(getdate(), 'yyyy-MM-dd') + ' ' + '" & vHoraEntrada & "')  + @DIFERENCA,108) AS HORA_SAIDA"
    
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "DECLARE @DIFERENCA AS DATETIME " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "SELECT @DIFERENCA =  " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "     CASE " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "         WHEN CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraSaida & "')) < 24 THEN " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "             CONVERT(datetime,'" & vHoraSaida & "') - CONVERT(datetime,'" & vHoraEntrada & "') " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "         ELSE " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "             REPLICATE('0', 2 - LEN(CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraSaida & "')) - CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraEntrada & "')))) + RTrim(CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraSaida & "')) - CONVERT(INT,CONVERT(VARCHAR(2),'" & vHoraEntrada & "'))) + ':00' " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "     END  " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & " " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "SELECT " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "   CONVERT(VARCHAR(10),FORMAT(CONVERT(DATETIME,FORMAT(getdate(), 'yyyy-MM-dd') + ' ' + '" & vHoraEntrada & "')  + @DIFERENCA,'yyyy-MM-dd'), 112) AS DATA_SAIDA, " & vbCrLf
    SqlAcertaHoraSaida = SqlAcertaHoraSaida & "   CONVERT(VARCHAR(5),CONVERT(DATETIME,FORMAT(getdate(), 'yyyy-MM-dd') + ' ' + '" & vHoraEntrada & "')  + @DIFERENCA,108) AS HORA_SAIDA"
    rsAcertaHoraSaida.Open SqlAcertaHoraSaida, cnBanco, adOpenKeyset, adLockReadOnly
    
    vDiferenca = rsAcertaHoraSaida.Fields(1)
    acertaHoraSaida = vDiferenca
    rsAcertaHoraSaida.Close
    
End Function

Private Sub exportaHorarioParaZeus(vChapa As String, vHEnt As String, vHSai As String, vCodColigada As Integer)
    Dim rsExportaHorarioParaZeus As New ADODB.Recordset
    Dim SqlExportaHorarioParaZeus  As String
    
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "UPDATE TBHORARIOS WITH (SERIALIZABLE) SET " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "   HORARIO_ENTRADA = '" & vHEnt & "', " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "   HORARIO_SAIDA = '" & vHSai & "' " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "WHERE " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "   CHAPA = '" & vChapa & "' AND " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "   CODCOLIGADA = 6 " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "IF @@rowcount = 0 " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "BEGIN " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "   INSERT INTO TBHORARIOS " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "       (CHAPA, HORARIO_ENTRADA, HORARIO_SAIDA, CODCOLIGADA) " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "   VALUES " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "       ('" & vChapa & "','" & vHEnt & "','" & vHSai & "'," & vCodColigada & ") " & vbCrLf
    SqlExportaHorarioParaZeus = SqlExportaHorarioParaZeus & "END"
    'cnBanco.Execute SqlExportaHorarioParaZeus, lRowsAffected, adExecuteNoRecords
    rsExportaHorarioParaZeus.Open SqlExportaHorarioParaZeus, cnBanco
End Sub


Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    If Conectar(Text7.Text, Text6.Text, Text5.Text, Text4.Text) = True Then
        carregaDadosConexaoRelogioPonto
        carregaDadosConexaoFlexJr
        Label1.Caption = "Importando batidas do relogio"
        importaBatidas RemoveMask(DTPicker1.Value), RemoveMask(DTPicker2.Value)
        'Timer2.Enabled = True
    Else
        MsgBox "Erro ao tentar realizar conexão com o servidor de banco de dados", vbCritical, "Atenção"
    End If
End Sub


